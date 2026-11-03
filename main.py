print("=== [ШАГ 1] СТАРТ ПРОГРАММЫ ===", flush=True)

import io
import os
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import docx
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from icalendar import Calendar, Event
from playwright.sync_api import sync_playwright

load_dotenv()

LOGIN = (os.getenv("PORTAL_LOGIN") or "10260472").strip()
PASSWORD = (os.getenv("PORTAL_PASSWORD") or "aS79XBqrVU").strip()
TARGET_GROUP = "56-11"
TARGET_GROUP_UUID = "b19cd1cf-f504-11f0-80c0-bc0d9b767ccf"

TG_TOKEN = (os.getenv("TG_TOKEN") or "").strip()
TG_CHAT_ID = (os.getenv("TG_CHAT_ID") or "").strip()

BASE_URL = "https://portal.petrocollege.ru"
URL = f"{BASE_URL}/schedule/schedule-and-substitutions/"
TIMEZONE = ZoneInfo("Europe/Moscow")

BELL_WEEKDAY = {
    1: ("09:00", "10:30"), 2: ("10:50", "12:20"), 3: ("13:05", "14:35"),
    4: ("14:50", "16:20"), 5: ("16:35", "18:05"), 6: ("18:15", "19:45"),
}
BELL_SATURDAY = {
    1: ("09:00", "10:30"), 2: ("10:45", "12:15"), 3: ("12:45", "14:15"),
    4: ("14:30", "16:00"), 5: ("16:15", "17:45"), 6: ("18:00", "19:30"),
}
MONTHS = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6,
    "июля": 7, "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
}

unique_events = {}
is_github = os.getenv("GITHUB_ACTIONS") == "true"

print(f"=== [ШАГ 2] ЗАПУСКАЕМ БРАУЗЕР (GitHub Actions: {is_github}) ===", flush=True)

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=is_github,
        args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"]
    )
    context = browser.new_context(
        viewport={"width": 1280, "height": 900},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    page = context.new_page()
    page.set_default_timeout(45000)

    print(f" -> Открываем сайт {URL} ...", flush=True)
    page.goto(URL, wait_until="commit", timeout=60000)
    page.wait_for_timeout(3000)

    if page.locator("input[type='password']").count() > 0:
        print(" -> Вводим логин и пароль...", flush=True)
        page.locator("input[name='USER_LOGIN'], input[name='login'], input[type='text']:not(.filter)").first.fill(LOGIN)
        page.locator("input[type='password']").first.fill(PASSWORD)
        btn = page.locator("input[type='submit'], button[type='submit'], input[name='Login']")
        if btn.count() > 0:
            btn.first.click()
        else:
            page.locator("input[type='password']").first.press("Enter")
        # Вместо зависающего networkidle ждем появления самого фильтра расписания
        page.wait_for_selector(".filter-group", timeout=45000)
    else:
        page.wait_for_selector(".filter-group", timeout=45000)

    print(" -> Фильтры успешно загружены!", flush=True)
    today = datetime.now(TIMEZONE)

    for shift in [0, 7]:
        target_dt = today + timedelta(days=shift)
        date_str = target_dt.strftime("%Y-%m-%d")
        print(f" -> Загружаем неделю ({date_str}) для группы {TARGET_GROUP}...", flush=True)

        # 1. Выставляем дату через JS (быстро и без зависаний)
        page.evaluate(
            """(d) => {
                const el = document.querySelector('input.filter-date');
                if (el) { el.value = d; el.dispatchEvent(new Event('change', {bubbles: true})); }
            }""",
            date_str
        )

        # 2. Выбираем "По группам"
        page.locator("select[name='filterType']").select_option("groups")
        page.locator("select[name='filterType']").dispatch_event("change")
        page.wait_for_timeout(2000)

        # 3. Устанавливаем группу 56-11 напрямую по её UUID
        page.evaluate(
        """({groupName, groupUuid}) => {
                const s = document.querySelector('#third-option');
                if (!s) return;
                const opt = Array.from(s.options).find(o => o.text.trim() === groupName);
                s.value = opt ? opt.value : groupUuid;
                s.dispatchEvent(new Event('change', { bubbles: true }));
                if (window.jQuery) {
                    window.jQuery('#third-option').val(s.value).trigger('change');
                }
            }""",
            {"groupName": TARGET_GROUP, "groupUuid": TARGET_GROUP_UUID}
        )
        page.wait_for_timeout(1000)

        # 4. Нажимаем "Показать" и ждем 4 секунды обновления карточек
        page.locator(".print-button").first.click()
        page.wait_for_timeout(4000)

        # 5. Собираем пары со страницы
        html = page.content()
        soup = BeautifulSoup(html, "html.parser")
        byear, bmonth = target_dt.year, target_dt.month

        for col in soup.select("#schedule-list .rasp-calendar-col"):
            d_el = col.select_one(".rasp-calendar__date")
            if not d_el:
                continue
            m_date = re.match(r"(\d{1,2})\s+([а-яА-ЯёЁ]+)", d_el.get_text(strip=True))
            if not m_date:
                continue
            day, month = int(m_date.group(1)), MONTHS.get(m_date.group(2).lower(), bmonth)
            year = byear + 1 if (bmonth == 12 and month == 1) else (byear - 1 if (bmonth == 1 and month == 12) else byear)
            d_key = f"{year:04d}-{month:02d}-{day:02d}"

            for row in col.select(".rasp-calendar__row"):
                disc_el = row.select_one(".rasp-calendar__discipline")
                if not disc_el:
                    continue
                raw = disc_el.get_text(strip=True)
                m_disc = re.match(r"^(\d+)\.\s*(.+)$", raw)
                p_num, subj = (int(m_disc.group(1)), m_disc.group(2).strip()) if m_disc else (1, raw)

                teacher, room = "Не указан", "Не указана"
                for el in row.select(".rasp-calendar__teacher-name"):
                    t = el.get_text(strip=True)
                    if t.lower().startswith("кабинет:"):
                        room = t.replace("Кабинет:", "").strip(" ,.")
                    else:
                        teacher = t
                sub_el = row.select_one(".rasp-calendar__address")
                sub_val = sub_el.get_text(strip=True) if sub_el else None

                unique_events[(d_key, p_num, sub_val)] = {
                    "date_key": d_key, "y": year, "m": month, "d": day, "num": p_num,
                    "subj": subj, "teacher": teacher, "room": room,
                    "sub": sub_val, "is_sub": False
                }

    cookies = {c["name"]: c["value"] for c in context.cookies()}
    browser.close()

events = list(unique_events.values())
print(f"=== [ШАГ 3] СОБРАНО ПАР: {len(events)}. ПРОВЕРЯЕМ ЗАМЕНЫ ===", flush=True)

for link in soup.select("#bx_886362991_add tbody tr a.table__icon-file")[:4]:
    d_txt, href = link.get_text(strip=True), link.get("href")
    if not href or not re.match(r"\d{2}\.\d{2}\.\d{4}", d_txt):
        continue
    d, m, y = map(int, d_txt.split("."))
    d_key = f"{y:04d}-{m:02d}-{d:02d}"
    try:
        r = requests.get(f"{BASE_URL}{href}", cookies=cookies, timeout=15)
        if r.ok:
            doc = docx.Document(io.BytesIO(r.content))
            for tbl in doc.tables:
                for row in tbl.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if TARGET_GROUP in " | ".join(cells):
                        p_num = next((int(re.search(r"\b([1-6])\b", c).group(1)) for c in cells[1:3] if re.search(r"\b([1-6])\b", c)), None)
                        info = " / ".join(dict.fromkeys([c for c in cells if c and c != TARGET_GROUP]))
                        print(f" -> Замена ({d_txt}): пара {p_num} -> {info}", flush=True)
                        for e in events:
                            if e["date_key"] == d_key and e["num"] == p_num:
                                e["subj"] = info
                                e["is_sub"] = True
    except Exception as err:
        print(f" -> Ошибка чтения файла замен {d_txt}: {err}", flush=True)

print("=== [ШАГ 4] СОХРАНЯЕМ КАЛЕНДАРЬ И ОТПРАВЛЯЕМ В TELEGRAM ===", flush=True)
cal = Calendar()
cal.add("prodid", "-//Petrocollege//RU")
cal.add("version", "2.0")
cal.add("x-wr-calname", f"Расписание {TARGET_GROUP}")
cal.add("x-wr-timezone", "Europe/Moscow")

for e in events:
    dt = datetime(e["y"], e["m"], e["d"])
    s_map = BELL_SATURDAY if dt.weekday() == 5 else BELL_WEEKDAY
    st, en = s_map.get(e["num"], ("18:15", "19:45"))
    if "45 МИНУТ" in e["subj"].upper() and e["num"] == 1:
        en = "09:45"
    sh, sm = map(int, st.split(":"))
    eh, em = map(int, en.split(":"))

    ev = Event()
    pref = "⚡ [ЗАМЕНА] " if e["is_sub"] else ""
    ev.add("summary", f"{pref}{e['subj']}" + (f" ({e['sub']})" if e["sub"] else ""))
    ev.add("dtstart", datetime(e["y"], e["m"], e["d"], sh, sm, tzinfo=TIMEZONE))
    ev.add("dtend", datetime(e["y"], e["m"], e["d"], eh, em, tzinfo=TIMEZONE))
    ev.add("location", f"Каб. {e['room']}")
    ev.add("description", f"Пара №: {e['num']}\nПреподаватель: {e['teacher']}\nКабинет: {e['room']}")
    ev.add("uid", f"{e['date_key']}-p{e['num']}-{e['sub'] or 'all'}@petrocollege.ru")
    cal.add_component(ev)

out_path = os.path.abspath("schedule.ics")
with open(out_path, "wb") as f:
    f.write(cal.to_ical())

if TG_TOKEN and TG_CHAT_ID:
    print(" -> Отправляем файл в Telegram...", flush=True)
    try:
        with open(out_path, "rb") as f:
            resp = requests.post(
                f"https://api.telegram.org/bot{TG_TOKEN}/sendDocument",
                data={"chat_id": TG_CHAT_ID, "caption": f"📅 Расписание {TARGET_GROUP} на 2 недели (пар: {len(events)})"},
                files={"document": f},
                timeout=20
            )
        print(f" -> Ответ Telegram ({resp.status_code}): {resp.text}", flush=True)
    except Exception as err:
        print(f" -> Ошибка отправки в Telegram: {err}", flush=True)
else:
    print(" -> [!] TG_TOKEN или TG_CHAT_ID не заданы в Secrets!", flush=True)

print(f"=== [ШАГ 5] ФИНИШ! Файл сохранен: {out_path} ===", flush=True)

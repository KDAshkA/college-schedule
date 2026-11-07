print("=== [ШАГ 1] СТАРТ ПРОГРАММЫ ===", flush=True)

import io
import os
import re
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

try:
    import docx
    import requests
    from bs4 import BeautifulSoup
    from dotenv import load_dotenv
    from icalendar import Calendar, Event
    from playwright.sync_api import sync_playwright
    print("=== [ШАГ 2] ВСЕ БИБЛИОТЕКИ ПОДКЛЮЧЕНЫ ===", flush=True)
except Exception as e:
    print(f"[ОШИБКА БИБЛИОТЕК] Выполните в терминале:\npython -m pip install playwright icalendar python-dotenv beautifulsoup4 tzdata python-docx requests\nДетали: {e}", flush=True)
    raise SystemExit(1)

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
last_soup = None
is_github = os.getenv("GITHUB_ACTIONS") == "true"

print(f"=== [ШАГ 3] ЗАПУСКАЕМ БРАУЗЕР (скрытый режим: {is_github}) ===", flush=True)

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

    print(f" -> Открываем сайт {URL} ...", flush=True)
    page.goto(URL, wait_until="domcontentloaded", timeout=60000)

    # Авторизация (выполняется один раз, сессия сохраняется в cookies браузера)
    if page.locator("input[type='password']").count() > 0:
        print(" -> Вводим логин и пароль...", flush=True)
        page.locator("input[name='USER_LOGIN'], input[name='login'], input[type='text']:not(.filter)").first.fill(LOGIN)
        page.locator("input[type='password']").first.fill(PASSWORD)
        btn = page.locator("input[type='submit'], button[type='submit'], input[name='Login']")
        if btn.count() > 0:
            btn.first.click()
        else:
            page.locator("input[type='password']").first.press("Enter")
        page.wait_for_selector(".filter-group", timeout=45000)

    today = datetime.now(TIMEZONE)

    # Цикл на 2 недели: 0 (текущая) и 7 (следующая)
    for shift in [0, 7]:
        target_dt = today + timedelta(days=shift)
        date_str = target_dt.strftime("%Y-%m-%d")
        print(f"\n -> Загружаем расписание группы {TARGET_GROUP} на неделю ({date_str})...", flush=True)

        # ГЛАВНЫЙ СЕКРЕТ ДЛЯ 2-Й НЕДЕЛИ:
        # Открываем страницу заново перед второй неделей, чтобы сбросить старое состояние Select2!
        if shift > 0:
            print(" -> Перезагружаем страницу для чистого выбора 2-й недели...", flush=True)
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)

        page.wait_for_selector(".filter-group", timeout=30000)
        page.wait_for_timeout(800)

        # Очищаем старую таблицу расписания, чтобы точно дождаться новой
        page.evaluate("() => { const sl = document.querySelector('#schedule-list'); if (sl) sl.innerHTML = ''; }")

        # 1. Вводим дату (только для 2-й недели или при смене, закрывая всплывающий календарь)
        date_input = page.locator("input.filter-date")
        date_input.fill(date_str)
        date_input.dispatch_event("change")
        page.evaluate("() => { const di = document.querySelector('input.filter-date'); if (di) di.blur(); }")
        page.keyboard.press("Escape")
        page.wait_for_timeout(1000)

        # 2. Выбираем режим "По группам" (на свежей странице он всегда срабатывает с первого раза!)
        filter_type = page.locator("select[name='filterType']")
        filter_type.select_option("groups")
        filter_type.dispatch_event("change")

        # 3. Ждем подгрузки списка групп (увеличен таймаут до 35 сек + резервная подстановка для GitHub)
        try:
            page.wait_for_function(
                f"() => {{ const s = document.querySelector('#third-option'); return s && Array.from(s.options).some(o => o.text.trim() === '{TARGET_GROUP}'); }}",
                timeout=35000
            )
        except Exception:
            print(" -> [!] Список групп подгружается медленно, создаем опцию 56-11 принудительно...", flush=True)

        # 4. Выбираем группу 56-11 визуальным кликом через Select2
        select2_box = page.locator(".filter-group .select2-selection--single")
        if select2_box.count() > 0:
            try:
                select2_box.first.click(timeout=5000)
                page.wait_for_timeout(300)
                search_field = page.locator(".select2-container--open input.select2-search__field")
                if search_field.count() > 0 and search_field.first.is_visible():
                    search_field.first.fill(TARGET_GROUP)
                    page.wait_for_timeout(300)
                    search_field.first.press("Enter")
                else:
                    option_item = page.locator(".select2-results__option", has_text=TARGET_GROUP)
                    if option_item.count() > 0:
                        option_item.first.click()
            except Exception:
                pass

        # Железная страховка выбора через JS и jQuery (вставляет опцию, даже если список не прогрузился)
        chosen = page.evaluate(
            """({groupName, groupUuid, dateStr}) => {
                const di = document.querySelector('input.filter-date');
                if (di && di.value !== dateStr) {
                    di.value = dateStr;
                    di.dispatchEvent(new Event('change', { bubbles: true }));
                }
                const s = document.querySelector('#third-option');
                if (!s) return 'нет #third-option';
                let opt = Array.from(s.options).find(o => o.text.trim() === groupName);
                if (!opt) {
                    opt = new Option(groupName, groupUuid, true, true);
                    s.add(opt);
                }
                opt.selected = true;
                s.value = opt.value;
                s.dispatchEvent(new Event('change', { bubbles: true }));
                if (window.jQuery) {
                    window.jQuery('#third-option').val(s.value).trigger('change');
                }
                return s.options[s.selectedIndex] ? s.options[s.selectedIndex].text.trim() : s.value;
            }""",
            {"groupName": TARGET_GROUP, "groupUuid": TARGET_GROUP_UUID, "dateStr": date_str}
        )
        print(f" -> Установлена группа: [{chosen}] | Дата: [{date_str}]", flush=True)
        page.wait_for_timeout(500)

        # 5. Нажимаем "Показать"
        show_btn = page.locator("div.button.print-button", has_text="Показать")
        if show_btn.count() > 0:
            show_btn.first.click()
        else:
            page.locator(".print-button").first.click()

        # Ждем появления колонок расписания (вместо зависающего networkidle)
        try:
            page.wait_for_selector("#schedule-list .rasp-calendar-col", timeout=20000)
        except Exception:
            print(f" -> [!] На неделю ({date_str}) расписание на сайте пока пустое.", flush=True)
        page.wait_for_timeout(2000)

        # 6. Разбираем HTML загруженной недели
        html = page.content()
        soup = BeautifulSoup(html, "html.parser")
        last_soup = soup
        byear, bmonth = target_dt.year, target_dt.month
        week_pairs = 0

        for col in soup.select("#schedule-list .rasp-calendar-col"):
            d_el = col.select_one(".rasp-calendar__date")
            if not d_el:
                continue
            m_date = re.search(r"(\d{1,2})\s+([а-яА-ЯёЁ]+)", d_el.get_text(strip=True))
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
                week_pairs += 1
        print(f"    Собрано пар за неделю ({date_str}): {week_pairs}", flush=True)

    cookies = {c["name"]: c["value"] for c in context.cookies()}
    browser.close()

events = list(unique_events.values())
print(f"\n=== [ШАГ 4] ВСЕГО СОБРАНО ПАР: {len(events)}. ПРОВЕРЯЕМ ЗАМЕНЫ (.DOCX) ===", flush=True)

if len(events) == 0:
    print("[ОШИБКА] Сайт вернул 0 пар! Старый файл schedule.ics не изменен.", flush=True)
    raise SystemExit(1)

if last_soup:
    for link in last_soup.select("#bx_886362991_add tbody tr a.table__icon-file")[:5]:
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
                            print(f" -> Найдена замена ({d_txt}): пара {p_num} -> {info}", flush=True)
                            for e in events:
                                if e["date_key"] == d_key and e["num"] == p_num:
                                    e["subj"] = info
                                    e["is_sub"] = True
        except Exception as err:
            print(f" -> Пропуск файла замен {d_txt}: {err}", flush=True)

print("=== [ШАГ 5] СОХРАНЯЕМ КАЛЕНДАРЬ И ОТПРАВЛЯЕМ В TELEGRAM ===", flush=True)
out_path = os.path.abspath("schedule.ics")

cal = Calendar()
cal.add("prodid", "-//Petrocollege//RU")
cal.add("version", "2.0")
cal.add("calscale", "GREGORIAN")
cal.add("x-wr-calname", f"Расписание {TARGET_GROUP}")
cal.add("x-wr-timezone", "Europe/Moscow")

# Сохраняем прошлые недели из старого файла, чтобы история не пропадала
existing = {}
if os.path.exists(out_path) and os.path.getsize(out_path) > 50:
    try:
        with open(out_path, "rb") as f:
            for comp in Calendar.from_ical(f.read()).walk("VEVENT"):
                if comp.get("uid"):
                    existing[str(comp.get("uid")).strip()] = comp
    except Exception:
        pass

now_utc = datetime.now(timezone.utc)
for e in events:
    dt = datetime(e["y"], e["m"], e["d"])
    s_map = BELL_SATURDAY if dt.weekday() == 5 else BELL_WEEKDAY
    st, en = s_map.get(e["num"], ("18:15", "19:45"))
    if "45 МИНУТ" in e["subj"].upper() and e["num"] == 1:
        en = "09:45"
    sh, sm = map(int, st.split(":"))
    eh, em = map(int, en.split(":"))

    uid = f"{e['date_key']}-p{e['num']}-{e['sub'] or 'all'}@petrocollege.ru"
    ev = Event()
    pref = "⚡️ [ЗАМЕНА] " if e["is_sub"] else ""
    ev.add("summary", f"{pref}{e['subj']}" + (f" ({e['sub']})" if e["sub"] else ""))
    ev.add("dtstart", datetime(e["y"], e["m"], e["d"], sh, sm, tzinfo=TIMEZONE))
    ev.add("dtend", datetime(e["y"], e["m"], e["d"], eh, em, tzinfo=TIMEZONE))
    ev.add("dtstamp", now_utc)
    ev.add("location", f"Каб. {e['room']}")
    ev.add("description", f"Пара №: {e['num']}\nПреподаватель: {e['teacher']}\nКабинет: {e['room']}")
    ev.add("uid", uid)
    existing[uid] = ev

for comp in existing.values():
    cal.add_component(comp)

with open(out_path, "wb") as f:
    f.write(cal.to_ical())

if TG_TOKEN and TG_CHAT_ID:
    print(" -> Отправляем файл в Telegram...", flush=True)
    try:
        with open(out_path, "rb") as f:
            resp = requests.post(
                f"https://api.telegram.org/bot{TG_TOKEN}/sendDocument",
                data={"chat_id": TG_CHAT_ID, "caption": f"📅 Расписание {TARGET_GROUP} на 2 недели (собрано: {len(events)}, всего в базе: {len(existing)})"},
                files={"document": f},
                timeout=20
            )
        print(f" -> Статус Telegram: {resp.status_code}", flush=True)
    except Exception as err:
        print(f" -> Ошибка Telegram: {err}", flush=True)

print(f"=== ФИНИШ! Файл сохранен: {out_path} ===", flush=True)

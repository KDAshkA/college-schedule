print("=== СТАРТ ПРОГРАММЫ ===")

import io
import os
import re
from datetime import datetime
from zoneinfo import ZoneInfo

try:
    import docx
    import requests
    from bs4 import BeautifulSoup
    from dotenv import load_dotenv
    from icalendar import Calendar, Event
    from playwright.sync_api import sync_playwright
    print("[1/5] Все библиотеки успешно подключены!")
except Exception as e:
    print(f"[ОШИБКА БИБЛИОТЕК] Выполните в терминале:\npython -m pip install playwright icalendar python-dotenv beautifulsoup4 tzdata python-docx requests\nДетали: {e}")
    raise SystemExit()

load_dotenv()

LOGIN = os.getenv("PORTAL_LOGIN") or "10260472"
PASSWORD = os.getenv("PORTAL_PASSWORD") or "aS79XBqrVU"
TARGET_GROUP = "56-11"
# Точный ID группы 56-11 из HTML-кода сайта (на случай резервного выбора)
TARGET_GROUP_UUID = "b19cd1cf-f504-11f0-80c0-bc0d9b767ccf"

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

print("[2/5] Открываем браузер...")
with sync_playwright() as p:
    # headless=True показывает окно браузера, чтобы вы видели выбор группы
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={"width": 1280, "height": 900})
    page = context.new_page()

    print(f"[3/5] Загружаем сайт {URL} ...")
    page.goto(URL, wait_until="domcontentloaded", timeout=60000)

    # Авторизация, если появилась форма входа
    if page.locator("input[type='password']").count() > 0:
        print(" -> Вводим логин и пароль...")
        page.locator("input[name='USER_LOGIN'], input[name='login'], input[type='text']:not(.filter)").first.fill(LOGIN)
        page.locator("input[type='password']").first.fill(PASSWORD)
        btn = page.locator("input[type='submit'], button[type='submit'], input[name='Login']")
        if btn.count() > 0:
            btn.first.click()
        else:
            page.locator("input[type='password']").first.press("Enter")
        page.wait_for_load_state("networkidle", timeout=30000)

    print(" -> Ждем появления блока фильтров...")
    page.wait_for_selector(".filter-group", timeout=30000)

    # 1. Выбираем фильтр "По группам" (value="groups")
    print(" -> Выбираем режим 'По группам'...")
    filter_type = page.locator("select[name='filterType']")
    filter_type.select_option("groups")
    filter_type.dispatch_event("change")

    # Ждем, пока в списке #third-option подгрузятся группы (в частности, группа 56-11)
    print(f" -> Ожидаем подгрузки списка групп и выбираем {TARGET_GROUP}...")
    page.wait_for_function(
        f"""() => {{
            const s = document.querySelector('#third-option');
            return s && Array.from(s.options).some(o => o.text.trim() === '{TARGET_GROUP}');
        }}""",
        timeout=15000
    )

    # 2. Выбираем группу 56-11 визуальным кликом через Select2 (как реальный пользователь)
    select2_box = page.locator(".filter-group .select2-selection--single")
    if select2_box.count() > 0:
        select2_box.first.click()
        page.wait_for_timeout(300)

        # Если есть поле поиска внутри выпадающего списка Select2 — печатаем туда "56-11"
        search_field = page.locator(".select2-container--open input.select2-search__field")
        if search_field.count() > 0 and search_field.first.is_visible():
            search_field.first.fill(TARGET_GROUP)
            page.wait_for_timeout(300)
            search_field.first.press("Enter")
        else:
            # Иначе кликаем прямо по пункту "56-11" в открывшемся списке
            option_item = page.locator(".select2-results__option", has_text=TARGET_GROUP)
            if option_item.count() > 0:
                option_item.first.click()

            # Страхуем выбор через JS (устанавливаем точный UUID группы 56-11 и вызываем событиеSelect2)
        page.evaluate(
            """({groupName, groupUuid}) => {
                const s = document.querySelector('#third-option');
                if (!s) return;
                const opt = Array.from(s.options).find(o => o.text.trim() === groupName);
                s.value = opt ? opt.value : groupUuid;
                s.dispatchEvent(new Event('change', {bubbles: true}));
                if (window.jQuery) {
                    window.jQuery('#third-option').val(s.value).trigger('change');
                }
            }""",
            {"groupName": TARGET_GROUP, "groupUuid": TARGET_GROUP_UUID}
        )

        page.wait_for_timeout(500)

        # 3. Нажимаем кнопку "Показать"
        print(" -> Нажимаем кнопку 'Показать' и ждем обновления таблицы...")
        show_btn = page.locator("div.button.print-button", has_text="Показать")
        if show_btn.count() > 0:
            show_btn.first.click()
        else:
            page.locator(".print-button").first.click()

        # Ждем завершения загрузки расписания
        page.wait_for_load_state("networkidle", timeout=20000)
        page.wait_for_timeout(2000)

        html = page.content()
        cookies = {c["name"]: c["value"] for c in context.cookies()}
        browser.close()

    print("[4/5] Читаем расписание группы 56-11 и проверяем файлы замен (.docx)...")
    soup = BeautifulSoup(html, "html.parser")
    events = []

    date_in = soup.select_one("input.filter-date")
    if date_in and date_in.get("value"):
        byear, bmonth = map(int, date_in["value"].split("-")[:2])
    else:
        byear, bmonth = datetime.now().year, datetime.now().month

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
            events.append({
                "date_key": d_key, "y": year, "m": month, "d": day, "num": p_num,
                "subj": subj, "teacher": teacher, "room": room,
                "sub": sub_el.get_text(strip=True) if sub_el else None, "is_sub": False
            })

    # Проверка таблицы замен (.docx)
    for link in soup.select("#bx_886362991_add tbody tr a.table__icon-file")[:3]:
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
                            p_num = next((int(re.search(r"\b([1-6])\b", c).group(1)) for c in cells[1:3] if
                                          re.search(r"\b([1-6])\b", c)), None)
                            info = " / ".join(dict.fromkeys([c for c in cells if c and c != TARGET_GROUP]))
                            print(f" -> Найдена замена в файле {d_txt}: пара {p_num} -> {info}")
                            matched = [e for e in events if e["date_key"] == d_key and e["num"] == p_num]
                            for e in matched:
                                e["subj"], e["is_sub"] = info, True
        except Exception as e:
            print(f" -> Не удалось прочитать файл замен {d_txt}: {e}")

        print("[5/5] Сохраняем файл schedule.ics ...")
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

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "schedule.ics")
        try:
            with open(desktop_path, "wb") as f:
                f.write(cal.to_ical())
        except Exception:
            pass

        print(f"=== ФИНИШ! Собрано пар: {len(events)}. Файл обновлен: {out_path} ===")
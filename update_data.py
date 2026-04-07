#!/usr/bin/env python3
import json, re, sys, io, urllib.request, urllib.parse
from datetime import datetime, date

YADISK_PUBLIC_KEY = "https://disk.360.yandex.ru/i/pbXk6jtc2OeG5Q"
INDEX_HTML = "index.html"

def download_from_yadisk():
    print("Получаю ссылку на скачивание с Яндекс Диска...")
    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=" + urllib.parse.quote(YADISK_PUBLIC_KEY)
    req = urllib.request.Request(api_url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.load(resp)
    download_url = data["href"]
    print("Скачиваю файл...")
    req2 = urllib.request.Request(download_url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req2, timeout=60) as resp:
        content = resp.read()
    print(f"Скачано: {len(content)} байт")
    return content


def classify_system(text):
    t = text.lower()
    if any(x in t for x in ['спринклер', 'орошени', 'огнезащит', 'огнестойкост', 'апс', 'аупт', 'аувп', 'ипр', ' пк ', 'пк №', 'впв', 'извещател', 'пожарн', 'эвакуац', 'дымов', 'мезонин', 'лду', 'противопожарн']):
        return 'Пожарная безопасность'
    if any(x in t for x in ['лвж', 'гж', 'гсм', 'аэрозол', 'топлив', 'горюч']):
        return 'ЛВЖ/ГСМ'
    if any(x in t for x in ['ворота', 'дверь', 'двери', 'замок', 'петли', 'докшелтер', 'аппарел', 'доквеллер', 'шлагбаум', 'жалюзи']):
        return 'Ворота и двери'
    if any(x in t for x in ['кабел', 'щит', 'силовой', 'зарядн', 'розетк', 'удлинител', 'электр', 'слаботочн']):
        return 'Электрика'
    if any(x in t for x in ['пол ', 'полов', 'топпинг', 'асфальт', 'покрыти']):
        return 'Покрытие полов'
    if any(x in t for x in ['стена', 'панел', 'перила', 'забор', 'ограждени', 'отбойник', 'бетон', 'конструкци', 'кровл', 'крыш', 'арматур', 'пандус']):
        return 'Конструктив'
    if any(x in t for x in ['вентиляц', 'вытяжк', 'приток', 'воздух']):
        return 'Вентиляция'
    if any(x in t for x in ['проживани', 'абч', 'бытовк', 'санузел', 'парковк', 'курилк', 'насажден']):
        return 'АБЧ/Территория'
    if any(x in t for x in ['стеллаж', 'склад', 'хранени', 'зона']):
        return 'Склад/Хранение'
    if any(x in t for x in ['план эвакуац', 'план', 'документ', 'журнал', 'инструкци', 'ответствен']):
        return 'Документация'
    return 'Прочее'

def parse_remarks(xlsx_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb["Акты, претензии СФН"]
    remarks = []
    current_act_num = None
    current_act_date = None
    current_object = None
    row_id = 0
    today = date.today()

    for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), 1):
        col0 = str(row[0]).strip() if row[0] else ""

        if re.search(r'Акт[уа]?\s*№', col0, re.IGNORECASE) and "2026" in col0:
            num_match = re.search(r'(ЧШ|ДД|ШШ|СПБ|МСК|00)-\d+', col0)
            current_act_num = num_match.group(0) if num_match else col0[:30]
            date_match = re.search(r'(\d+)\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})', col0)
            months = {"января":1,"февраля":2,"марта":3,"апреля":4,"мая":5,"июня":6,
                      "июля":7,"августа":8,"сентября":9,"октября":10,"ноября":11,"декабря":12}
            if date_match:
                d, m, y = date_match.groups()
                current_act_date = f"{y}-{months[m]:02d}-{int(d):02d}"
            else:
                current_act_date = "2026-01-01"
            current_object = None
            continue

        if not current_act_num:
            continue

        text = col0.replace('\xa0', '').strip()
        if not text or text in ("None", ""):
            continue
        if re.search(r'Акт[уа]?\s*№', text, re.IGNORECASE):
            if "2026" not in text:
                current_act_num = None
            continue

        obj = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        if obj and obj != "None":
            current_object = obj
        responsible = str(row[5]).strip() if len(row) > 5 and row[5] else ""

        deadline = ""
        for col_idx in [8, 7, 6]:
            if len(row) > col_idx and row[col_idx]:
                d = row[col_idx]
                if isinstance(d, (datetime, date)):
                    deadline = d.strftime("%Y-%m-%d") if isinstance(d, datetime) else d.isoformat()
                elif "2026" in str(d) or "2025" in str(d):
                    deadline = str(d)[:10]
                if deadline:
                    break

        comments = []
        for ci in [9, 10, 11]:
            if len(row) > ci and row[ci] and str(row[ci]) not in ("None", ""):
                comments.append(str(row[ci]).strip())
        full_comment = " | ".join(filter(None, comments))
        comment_lower = full_comment.lower()
        text_lower = text.lower()

        if any(x in comment_lower for x in ["устранено", "выполнено", "готово", "убрали", "демонтир"]):
            status = "done"
        elif any(x in comment_lower for x in ["направлено", "мониторинг", "регулярн", "письмо"]) or re.search(r'2026-\d{2}-\d{2}', full_comment):
            status = "progress"
        else:
            status = "open"

        if status != "done" and deadline:
            try:
                dl = datetime.strptime(deadline[:10], "%Y-%m-%d").date()
                if dl < today:
                    status = "overdue"
            except:
                pass

        eviction_risk = any(x in text_lower for x in [
            "мезонин", "огнезащит", "огнестойкост", "лвж", "гж", "гсм",
            "аэрозольн", "проживание", "система орошени", "орошени"
        ])
        fine_risk = any(x in text_lower for x in ["штраф", "нарушен", "предписани", "протокол"]) or eviction_risk
        is_rvb = "рвб" in comment_lower or (len(comments) > 0 and "РВБ" in comments[0])

        risk = "eviction" if eviction_risk else ("fine" if fine_risk else "")
        resp = "rvb" if is_rvb else ""

        row_id += 1
        remarks.append({
            "id": row_id,
            "object": current_object or "",
            "actNum": current_act_num,
            "actDate": current_act_date,
            "text": text,
            "responsible": responsible,
            "resp": resp,
            "respName": responsible,
            "deadline": deadline,
            "comment": full_comment,
            "status": status,
            "risk": risk,
            "evictionRisk": eviction_risk,
            "fineRisk": fine_risk,
            "isRvb": is_rvb,
            "system": classify_system(text)
        })

    print(f"Распарсено {len(remarks)} замечаний")
    evict_open = sum(1 for r in remarks if r['evictionRisk'] and r['status'] != 'done')
    print(f"Риск выселения (открытых): {evict_open}")
    return remarks

def update_index_html(remarks):
    with open(INDEX_HTML, encoding="utf-8") as f:
        html = f.read()
    new_json = json.dumps(remarks, ensure_ascii=False, separators=(",", ":"))
    new_data = f'const STATIC_DATA = {new_json};'

    marker_start = "const STATIC_DATA = ["
    marker_end = "];"
    idx_start = html.find(marker_start)
    if idx_start == -1:
        print("ОШИБКА: STATIC_DATA не найден в index.html")
        return False
    idx_end = html.find(marker_end, idx_start)
    if idx_end == -1:
        print("ОШИБКА: конец STATIC_DATA не найден")
        return False
    new_html = html[:idx_start] + new_data + html[idx_end + len(marker_end):]

    with open(INDEX_HTML, "w", encoding="utf-8") as f:
        f.write(new_html)
    print(f"index.html обновлён ({len(remarks)} замечаний)")
    return True

if __name__ == "__main__":
    try:
        xlsx_bytes = download_from_yadisk()
        remarks = parse_remarks(xlsx_bytes)
        if not remarks:
            print("Замечания не найдены")
            sys.exit(1)
        if not update_index_html(remarks):
            sys.exit(1)
        print("Готово!")
    except Exception as e:
        print(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

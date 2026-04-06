#!/usr/bin/env python3
"""
GitHub Actions скрипт: скачивает Excel с Яндекс Диска,
парсит замечания 2026 года и обновляет STATIC_DATA в index.html
"""

import json
import re
import urllib.request
import urllib.parse
import urllib.error
import sys
import io
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

def parse_remarks(xlsx_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    ws = wb["Акты, претензии СФН"]
    remarks = []
    current_act_num = None
    current_act_date = None
    current_object = None
    row_id = 0
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row_idx < 315:
            continue
        col0 = str(row[0]).strip() if row[0] else ""
        if re.match(r'^(ЧШ|ДД)-\d{8}$', col0):
            current_act_num = col0
            if row[1]:
                try:
                    if isinstance(row[1], (datetime, date)):
                        current_act_date = row[1].strftime("%Y-%m-%d")
                    else:
                        current_act_date = str(row[1])[:10]
                except:
                    current_act_date = str(row[1])
            current_object = str(row[2]).strip() if row[2] else ""
            continue
        if not col0 or col0 in ("None", "") or not current_act_num:
            continue
        if any(x in col0.lower() for x in ["замечание", "наименование", "пункт"]):
            continue
        text = col0
        risk = str(row[2]).strip() if len(row) > 2 and row[2] and str(row[2]) != "None" else ""
        deadline1 = str(row[6]).strip() if len(row) > 6 and row[6] else ""
        deadline2 = str(row[7]).strip() if len(row) > 7 and row[7] else ""
        comment = str(row[9]).strip() if len(row) > 9 and row[9] else ""
        comment2 = str(row[10]).strip() if len(row) > 10 and row[10] else ""
        comment3 = str(row[11]).strip() if len(row) > 11 and row[11] else ""
        full_comment = " | ".join(filter(None, [comment, comment2, comment3]))
        responsible = str(row[5]).strip() if len(row) > 5 and row[5] else ""
        deadline = ""
        for d in [deadline2, deadline1]:
            if d and d not in ("None", "") and len(d) >= 8:
                deadline = d[:10]
                break
        comment_lower = full_comment.lower()
        today = date.today()
        if any(x in comment_lower for x in ["устранено", "выполнено", "готово", "убрали", "демонтир"]):
            status = "done"
        elif any(x in comment_lower for x in ["направлено", "мониторинг", "регулярн", "письмо"]):
            status = "progress"
        elif re.search(r'2026', full_comment):
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
        text_lower = text.lower()
        eviction_risk = any(x in text_lower for x in ["мезонин", "огнезащит", "огнестойкост", "лвж", "гж", "гсм", "аэрозольн", "проживание", "система орошени"])
        fine_risk = any(x in text_lower for x in ["штраф", "нарушен", "предписани", "протокол"]) or eviction_risk
        is_rvb = "рвб" in comment_lower or "rvb" in comment_lower
        row_id += 1
        remarks.append({
            "id": row_id,
            "object": current_object,
            "actNum": current_act_num,
            "actDate": current_act_date or "",
            "text": text,
            "risk": risk,
            "responsible": responsible,
            "deadline": deadline,
            "comment": full_comment,
            "status": status,
            "evictionRisk": eviction_risk,
            "fineRisk": fine_risk,
            "isRvb": is_rvb
        })
    print(f"Распарсено {len(remarks)} замечаний")
    return remarks

def update_index_html(remarks):
    with open(INDEX_HTML, encoding="utf-8") as f:
        html = f.read()
    new_json = json.dumps(remarks, ensure_ascii=False, separators=(",", ":"))
    old_pattern = r'const STATIC_DATA = \[.*?\];'
    new_data = f'const STATIC_DATA = {new_json};'
    new_html = re.sub(old_pattern, new_data, html, flags=re.DOTALL)
    if new_html == html:
        print("ОШИБКА: STATIC_DATA не найден в index.html")
        return False
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
        success = update_index_html(remarks)
        if not success:
            sys.exit(1)
        print("Готово!")
    except Exception as e:
        print(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

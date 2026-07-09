#!/usr/bin/env python3
"""
TPay Loyalty – Excel → data.json / data2.json + draw_data.json converter
=========================================================================
Хэрэглээ / Usage:
  1-р сарын аян (Шинэ зээлийн):
    python3 convert.py month1_data.xlsx

  2-р сарын аян (Эргэн төлөлтийн):
    python3 convert.py month2_data.xlsx --month 2

  3-р сарын аян (Давтан зээлийн):
    python3 convert.py month3_data.xlsx --month 3

Excel файлын формат (5 багана):
  A: Харилцагч нэр      (жш: САРАНЦЭЦЭГ)
  B: Харилцагч овог     (жш: Энхтөр)
  C: Харилцагч утас     (жш: 96449901)
  D: Киоск нэр          (жш: KIOSK 32)
  E: Зээл авсан огноо   (жш: 2026-07-10)

Гаралт:
  --month 1 (default) → data.json  + draw_data.json
  --month 2           → data2.json + draw_data2.json
  --month 3           → data3.json + draw_data3.json
"""

import sys
import json
import os
import re
from datetime import date, datetime

try:
    import openpyxl
except ImportError:
    print("openpyxl олдсонгүй. Автоматаар суулгаж байна...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl
    print("✅ openpyxl суулгагдлаа.\n")


def clean_phone(raw):
    if raw is None:
        return None
    phone = str(raw).strip().replace(" ", "").replace("-", "")
    # Remove leading +976 or 976
    if phone.startswith("+976"):
        phone = phone[4:]
    elif phone.startswith("976") and len(phone) == 11:
        phone = phone[3:]
    if not phone.isdigit() or len(phone) != 8:
        return None
    return phone


def fmt_date(raw):
    if raw is None:
        return ""
    if isinstance(raw, (datetime, date)):
        return raw.strftime("%Y-%m-%d")
    s = str(raw).strip()
    # try common formats
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s[:len(fmt)], fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return s[:10]


def detect_format(headers):
    """
    'tpay' — TPay системийн шууд экспорт: Төрөл | Регистр | Овог | Нэр | Утас
    'new'  — гараар бэлдсэн 5 багана:     Нэр | Овог | Утас | Киоск | Огноо
    'old'  — хуучин 2 багана:              Утас | Тикет тоо
    """
    h = [str(x).strip().lower() if x else "" for x in headers]
    if any("төрөл" in x for x in h):
        return "tpay"
    if any("нэр" in x or "овог" in x for x in h):
        return "new"
    return "old"


def convert(excel_path: str, month: int = 1) -> None:
    suffix = "" if month == 1 else str(month)
    print(f"\n📋 {month}-р сарын аян → data{suffix}.json / draw_data{suffix}.json\n")
    if not os.path.exists(excel_path):
        print(f"Файл олдсонгүй: {excel_path}")
        sys.exit(1)

    base_dir = os.path.dirname(os.path.abspath(excel_path))

    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    except Exception as e:
        print(f"\n❌  Excel файл унших боломжгүй: {e}")
        sys.exit(1)

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        print("❌ Хоосон файл.")
        sys.exit(1)

    headers = rows[0]
    fmt = detect_format(headers)
    fmt_labels = {
        'tpay': 'TPay экспорт (Төрөл|Регистр|Овог|Нэр|Утас) — Сунгалт=1эрх, Хаасан=2эрх',
        'new':  '5 багана (Нэр|Овог|Утас|Киоск|Огноо)',
        'old':  '2 багана (Утас|Тикет тоо)',
    }
    print(f"   Формат илэрлээ: {fmt_labels.get(fmt, fmt)}")

    # ── Build entries ──────────────────────────────────────
    entries = []   # for draw_data.json  — one entry per ticket row
    phone_counts = {}  # for data.json   — {phone: ticket_count}
    skipped = 0

    for row_idx, row in enumerate(rows[1:], start=2):
        if not any(row):
            continue  # skip fully empty rows

        if fmt == "tpay":
            # Columns: Төрөл | Регистр | Овог | Нэр | Утасны дугаар
            töröl   = str(row[0]).strip() if row[0] else ""
            surname = str(row[2]).strip() if row[2] else ""
            name    = str(row[3]).strip() if row[3] else ""
            phone   = clean_phone(row[4])
            if töröl == "Сунгалт":
                tickets = 1
            elif töröl in ("Хаасан", "Хаах", "Хаасан зээл"):
                tickets = 2
            else:
                skipped += 1
                continue
        elif fmt == "new":
            name    = str(row[0]).strip() if row[0] else ""
            surname = str(row[1]).strip() if row[1] else ""
            phone   = clean_phone(row[2])
            kiosk   = str(row[3]).strip() if row[3] else ""
            loan_dt = fmt_date(row[4])
            tickets = 1
        else:
            phone   = clean_phone(row[0])
            name    = ""
            surname = ""
            kiosk   = ""
            loan_dt = ""
            try:
                tickets = int(float(str(row[1]))) if row[1] is not None else 1
            except Exception:
                tickets = 1

        if phone is None:
            skipped += 1
            continue

        if fmt == "tpay":
            for _ in range(tickets):
                entries.append({
                    "phone":   phone,
                    "name":    name,
                    "surname": surname,
                    "kiosk":   "",
                    "date":    ""
                })
            phone_counts[phone] = phone_counts.get(phone, 0) + tickets
        elif fmt == "new":
            entries.append({
                "phone":   phone,
                "name":    name,
                "surname": surname,
                "kiosk":   kiosk,
                "date":    loan_dt
            })
            phone_counts[phone] = phone_counts.get(phone, 0) + 1
        else:
            for _ in range(max(tickets, 1)):
                entries.append({"phone": phone, "name": "", "surname": "", "kiosk": "", "date": ""})
            phone_counts[phone] = phone_counts.get(phone, 0) + max(tickets, 1)

    wb.close()

    total_tickets    = len(entries)
    unique_phones    = len(phone_counts)
    today            = str(date.today())

    # ── Write dataN.json (for index.html lookup + stats) ──
    data_json = {
        "_updated": today,
        "_note": f"convert.py --month {month} ашиглан үүсгэсэн.",
    }
    data_json.update(phone_counts)

    data_json_path = os.path.join(base_dir, f"data{suffix}.json")
    with open(data_json_path, "w", encoding="utf-8") as f:
        json.dump(data_json, f, ensure_ascii=False, indent=2)

    # ── Write draw_dataN.json (for draw.html) ─────────────
    draw_json = {
        "_updated": today,
        "_total_tickets": total_tickets,
        "_unique_phones": unique_phones,
        "entries": entries
    }
    draw_json_path = os.path.join(base_dir, f"draw_data{suffix}.json")
    with open(draw_json_path, "w", encoding="utf-8") as f:
        json.dump(draw_json, f, ensure_ascii=False, indent=2)

    print(f"✅  Амжилттай үүслээ")
    print(f"   Нийт тасалбар:     {total_tickets}")
    print(f"   Өвөрмөц харилцагч: {unique_phones}")
    print(f"   Алгасав:           {skipped}")
    print(f"   {data_json_path}")
    print(f"   {draw_json_path}")
    print()


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    month_arg = 1
    for i, a in enumerate(sys.argv[1:]):
        if a == "--month" and i + 2 < len(sys.argv):
            try:
                month_arg = int(sys.argv[i + 2])
            except ValueError:
                pass
    if not args:
        print(__doc__)
        sys.exit(0)
    convert(args[0], month=month_arg)

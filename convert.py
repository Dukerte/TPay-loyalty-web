#!/usr/bin/env python3
"""
TPay Loyalty – Excel → data.json converter
==========================================
Хэрэглээ / Usage:
  python3 convert.py raffle_data.xlsx

Excel файлын формат (2 багана):
  A: Бүртгэлтэй утасны дугаар   (жш: 99001234)
  B: Сугалааны эрх               (жш: 3)

Гаралт: data.json  (index.html-тэй ижил хавтсанд)
"""

import sys
import json
import os
from datetime import date

try:
    import openpyxl
except ImportError:
    print("openpyxl олдсонгүй. Автоматаар суулгаж байна...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl
    print("✅ openpyxl суулгагдлаа.\n")


def iter_rows(file_path):
    """Yield (phone_raw, tickets_raw) from .xlsx or .csv"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        import csv
        with open(file_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            next(reader, None)  # skip header
            for row in reader:
                yield row[0] if len(row) > 0 else None, row[1] if len(row) > 1 else None
    else:
        try:
            wb = openpyxl.load_workbook(file_path)
        except Exception as e:
            print(f"\n❌  Excel файл унших боломжгүй: {e}")
            print("   Шийдэл: Excel дээр файлаа нээж, 'Save As' → raffle_data.xlsx гэж хадгалаад дахин оролдно уу.")
            print("   Эсвэл CSV форматаар хадгалж:  python3 convert.py raffle_data.csv\n")
            sys.exit(1)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            yield row[0] if len(row) > 0 else None, row[1] if len(row) > 1 else None


def convert(excel_path: str, output_path: str = "data.json") -> None:
    if not os.path.exists(excel_path):
        print(f"Файл олдсонгүй: {excel_path}")
        sys.exit(1)

    data: dict = {
        "_updated": str(date.today()),
        "_note": "Энэ файлыг convert.py ашиглан Excel-ээс үүсгэсэн.",
    }

    skipped = 0
    added = 0

    for row_idx, (phone_raw, tickets_raw) in enumerate(iter_rows(excel_path), start=2):
        if phone_raw is None:
            continue

        phone = str(phone_raw).strip().replace(" ", "").replace("-", "")

        # 8 оронтой тоо байх ёстой
        if not phone.isdigit() or len(phone) != 8:
            print(f"  [мөр {row_idx}] Алгасав – буруу дугаар: '{phone_raw}'")
            skipped += 1
            continue

        try:
            tickets = int(float(str(tickets_raw))) if tickets_raw is not None else 0
        except (ValueError, TypeError):
            tickets = 0

        data[phone] = tickets
        added += 1

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n✅  data.json амжилттай үүслээ")
    print(f"   Нийт бүртгэл: {added}  |  Алгасав: {skipped}")
    print(f"   Хадгалагдсан: {os.path.abspath(output_path)}")

    # Also patch the inline fallback in index.html so file:// works too
    html_path = os.path.join(os.path.dirname(os.path.abspath(output_path)), "index.html")
    if os.path.exists(html_path):
        phone_only = {k: v for k, v in data.items() if not k.startswith("_")}
        inline_json = json.dumps(phone_only, ensure_ascii=False)
        with open(html_path, "r", encoding="utf-8") as f:
            html = f.read()
        import re
        patched = re.sub(
            r'(let raffleData = )\{[^}]*\};',
            f'let raffleData = {inline_json};',
            html
        )
        if patched != html:
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(patched)
            print(f"   index.html inline дата шинэчлэгдлээ ✓")
        else:
            print(f"   index.html: inline дата олдсонгүй, алгасав")
    print()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)
    excel_file = sys.argv[1]
    out_file = sys.argv[2] if len(sys.argv) > 2 else "data.json"
    convert(excel_file, out_file)

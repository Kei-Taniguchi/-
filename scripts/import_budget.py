from pathlib import Path
import json
import re
from datetime import datetime

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "家計.xlsx"
OUT = ROOT / "data" / "budget-data.json"


def value(v):
    if v is None:
        return None
    if hasattr(v, "isoformat"):
        return v.isoformat()
    if isinstance(v, (str, int, float, bool)):
        return v
    return str(v)


def row_values(ws):
    for row in ws.iter_rows(values_only=True):
        yield [value(v) for v in row]


def main():
    if not BOOK.exists():
        raise SystemExit(f"{BOOK.name} が見つかりません")

    wb = openpyxl.load_workbook(BOOK, data_only=True, read_only=True)
    sheets = {}
    for ws in wb.worksheets:
        sheets[ws.title] = {
            "title": ws.title,
            "rows": list(row_values(ws)),
        }

    # Keep the raw sheet data so future spreadsheet layouts can be mapped
    # without requiring another conversion step. Also expose likely budget
    # categories detected from sheet names/cell text.
    categories = {"リボ": [], "分割": [], "固定費": [], "電気": [], "ガス": [], "水道": []}
    for title, sheet in sheets.items():
        text = " ".join(str(x) for row in sheet["rows"] for x in row if x is not None)
        for category in categories:
            if category in title or category in text:
                categories[category].append(title)

    result = {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "source": BOOK.name,
        "sheetNames": list(wb.sheetnames),
        "categories": categories,
        "sheets": sheets,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Generated {OUT} from {len(sheets)} sheets")


if __name__ == "__main__":
    main()

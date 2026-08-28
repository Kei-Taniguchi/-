from pathlib import Path
import json
import re
from collections import OrderedDict
from datetime import datetime

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "家計.xlsx"
OUT = ROOT / "data" / "budget-data.json"
MONTH_RE = re.compile(r"^(\d{4})年(\d{1,2})月分$")
INSTALLMENT_RE = re.compile(r"\d+\s*/\s*\d+")
END_RE = re.compile(r"～\s*\d{4}\s*/\s*\d+")
FIXED_KEYWORDS = ("家賃", "LIFELINE", "通信", "携帯", "スマホ", "保険", "サブスク")

def value(v):
    if v is None: return None
    if hasattr(v, "isoformat"): return v.isoformat()
    if isinstance(v, (str, int, float, bool)): return v
    return str(v)

def number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)

def row_values(ws):
    for row in ws.iter_rows(values_only=True): yield [value(v) for v in row]

def parse_ratio(v):
    if v is None: return None
    m = re.search(r"(\d+)\s*/\s*(\d+)", str(v))
    return (int(m.group(1)), int(m.group(2))) if m else None

def classify(detail, h, i):
    text = " ".join(str(x) for x in (detail, h, i) if x is not None)
    if INSTALLMENT_RE.search(text) or END_RE.search(text): return "分割"
    if detail and (str(detail).strip().upper() == "R" or "リボ" in str(detail)): return "リボ"
    if detail and "電気" in str(detail): return "電気"
    if detail and "ガス" in str(detail): return "ガス"
    if detail and "水道" in str(detail): return "水道"
    if detail and any(k.lower() in str(detail).lower() for k in FIXED_KEYWORDS): return "固定費"
    return "その他"

def parse_month_sheet(ws):
    income = 0; payments = []
    for row_no, row in enumerate(ws.iter_rows(values_only=True), start=1):
        a = row[0] if len(row) > 0 else None; b = row[1] if len(row) > 1 else None
        e = row[4] if len(row) > 4 else None; g = row[6] if len(row) > 6 else None
        h = row[7] if len(row) > 7 else None; i = row[8] if len(row) > 8 else None
        if row_no != 7:
            if number(a): income += a
            if number(b): income += b
        if e is None or not number(g) or g == 0: continue
        ratio = parse_ratio(h) or parse_ratio(i)
        payments.append({"detail": str(e), "amount": g, "category": classify(e, h, i), "row": row_no, "progress": {"paid": ratio[0], "total": ratio[1]} if ratio else None, "schedule": str(i) if i is not None else None, "h": str(h) if h is not None else None, "i": str(i) if i is not None else None})
    balance = ws["B7"].value
    return income, balance if number(balance) else None, payments

def main():
    if not BOOK.exists(): raise SystemExit(f"{BOOK.name} が見つかりません")
    wb = openpyxl.load_workbook(BOOK, data_only=True, read_only=True)
    sheets = {}; months = []; payment_index = OrderedDict()
    category_sheets = {"リボ": [], "分割": [], "固定費": [], "電気": [], "ガス": [], "水道": []}
    for ws in wb.worksheets:
        rows = list(row_values(ws)); sheets[ws.title] = {"title": ws.title, "rows": rows}
        m = MONTH_RE.match(ws.title)
        if not m: continue
        month = f"{m.group(1)}-{int(m.group(2)):02d}"; income, balance, payments = parse_month_sheet(ws)
        for p in payments:
            key = (p["detail"], p["category"])
            item = payment_index.setdefault(key, {"name": p["detail"], "type": p["category"], "monthly": p["amount"], "latestMonth": month, "progress": p["progress"], "end": p["schedule"], "history": []})
            item["monthly"] = p["amount"]; item["latestMonth"] = month
            if p["progress"]: item["progress"] = p["progress"]
            if p["schedule"]: item["end"] = p["schedule"]
            item["history"].append({"month": month, "amount": p["amount"], "row": p["row"]})
            if p["category"] in category_sheets and ws.title not in category_sheets[p["category"]]: category_sheets[p["category"]].append(ws.title)
        months.append({"month": month, "sheet": ws.title, "income": income, "balance": balance, "payments": payments, "paymentTotal": sum(p["amount"] for p in payments)})
    payment_items = []
    for item in payment_index.values():
        progress = item.get("progress"); remaining_count = remaining_amount = None
        if progress:
            remaining_count = max(0, progress["total"] - progress["paid"]); remaining_amount = item["monthly"] * remaining_count
        payment_items.append({**item, "remainingCount": remaining_count, "remainingAmount": remaining_amount})
    result = {"generatedAt": datetime.now().isoformat(timespec="seconds"), "source": BOOK.name, "rules": {"incomeColumns": ["A", "B"], "balanceCell": "B7", "paymentDetailColumn": "E", "paymentAmountColumn": "G", "installmentColumns": ["H", "I"]}, "sheetNames": list(wb.sheetnames), "categories": category_sheets, "months": sorted(months, key=lambda x: x["month"]), "paymentItems": payment_items, "sheets": sheets}
    OUT.parent.mkdir(parents=True, exist_ok=True); OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Generated {OUT} from {len(sheets)} sheets; monthly sheets={len(months)}; payment items={len(payment_items)}")

if __name__ == "__main__": main()

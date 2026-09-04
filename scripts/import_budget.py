from pathlib import Path
import json
import re
from datetime import datetime

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
BOOK_XLSM = ROOT / "家計.xlsm"
BOOK_XLSX = ROOT / "家計.xlsx"
OUT = ROOT / "data" / "budget-data.json"
MONTH_RE = re.compile(r"^(\d{4})年(\d{1,2})月分$")
RATIO_RE = re.compile(r"(\d+)\s*/\s*(\d+)")
END_RE = re.compile(r"[～~]\s*(\d{4})\s*/\s*(\d+)")


def number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def parse_ratio(*values):
    for v in values:
        if v is None:
            continue
        m = RATIO_RE.search(str(v))
        if m:
            return int(m.group(1)), int(m.group(2))
    return None


def is_completed_gray_row(ws, row_no):
    # The workbook uses the light-gray theme tint on the completed rows.
    # Require the gray style on at least three cells in B:F to avoid treating
    # ordinary formatting as a completed payment.
    hits = 0
    for col in range(2, 7):
        color = ws.cell(row_no, col).fill.fgColor
        if color.type == "theme" and color.theme == 0 and color.tint < -0.20:
            hits += 1
    return hits >= 3


def classify_payment(detail, h, i):
    text = " ".join(str(x) for x in (detail, h, i) if x is not None)
    if str(detail).strip().upper() == "R" or "リボ" in str(detail):
        return "リボ"
    if RATIO_RE.search(text) or END_RE.search(text):
        return "分割"
    return None


def parse_month_sheet(ws, month):
    income = 0
    payments = []
    # From 2026-07 onward A:C are pre-expense-settlement figures and must not
    # be counted as income. Before that, preserve the original A/B income rule.
    if month < "2026-07":
        for row_no in range(1, ws.max_row + 1):
            if row_no == 7:
                continue
            for col in (1, 2):
                v = ws.cell(row_no, col).value
                if number(v):
                    income += v

    for row_no in range(1, ws.max_row + 1):
        detail = ws.cell(row_no, 5).value
        amount = ws.cell(row_no, 7).value
        h = ws.cell(row_no, 8).value
        i = ws.cell(row_no, 9).value
        if detail is None or not number(amount) or amount == 0:
            continue
        category = classify_payment(detail, h, i)
        if not category:
            continue
        ratio = parse_ratio(h, i)
        payments.append({
            "name": str(detail),
            "type": category,
            "amount": amount,
            "progress": list(ratio) if ratio else None,
            "row": row_no,
        })

    balance = ws["B7"].value
    return income, balance if number(balance) else None, payments


def extract_card_schedules(ws):
    schedules = []
    for row_no in range(1, ws.max_row + 1):
        name = ws.cell(row_no, 9).value
        due = ws.cell(row_no, 10).value
        closing = ws.cell(row_no, 11).value
        if name is None:
            continue
        schedules.append({
            "name": str(name),
            "引落": str(due) if due is not None else "",
            "締め日": str(closing) if closing is not None else "",
        })
    return schedules


def extract_split_items(ws):
    items = []
    for row_no in range(3, ws.max_row + 1):
        name = ws.cell(row_no, 2).value
        if name is None:
            continue
        if is_completed_gray_row(ws, row_no):
            continue
        total = ws.cell(row_no, 3).value
        monthly = ws.cell(row_no, 4).value
        end = ws.cell(row_no, 5).value
        remaining = ws.cell(row_no, 6).value
        if not number(monthly):
            continue
        item_type = "リボ" if str(name).strip().upper() == "R" else "分割"
        items.append({
            "name": str(name),
            "type": item_type,
            "total": total if number(total) else None,
            "monthly": monthly,
            "end": str(end) if end is not None else None,
            "remainingAmount": remaining if number(remaining) else None,
            "row": row_no,
        })
    return items


def extract_life_costs(months):
    result = []
    for m in months:
        costs = {"家賃": 0, "ガス": 0, "電気": 0, "携帯": 0, "水道": 0}
        for p in m["raw_payments"]:
            text = p["name"]
            for key in costs:
                if key in text:
                    costs[key] += p["amount"]
        result.append({"month": m["month"], **costs})
    return result


def main():
    book = BOOK_XLSM if BOOK_XLSM.exists() else BOOK_XLSX
    if not book.exists():
        raise SystemExit("家計.xlsm または 家計.xlsx が見つかりません")

    wb = openpyxl.load_workbook(book, data_only=True, read_only=False, keep_vba=book.suffix.lower() == ".xlsm")
    months = []
    split_sheet = wb["分割分"] if "分割分" in wb.sheetnames else None

    for ws in wb.worksheets:
        m = MONTH_RE.match(ws.title)
        if not m:
            continue
        month = f"{m.group(1)}-{int(m.group(2)):02d}"
        income, balance, payments = parse_month_sheet(ws, month)
        raw_life = [{"name": p["name"], "amount": p["amount"]} for p in payments]
        months.append({
            "month": month,
            "sheet": ws.title,
            "income": income,
            "balance": balance,
            "payments": payments,
            "paymentTotal": sum(p["amount"] for p in payments),
            "raw_payments": raw_life,
        })

    months.sort(key=lambda x: x["month"])
    life_cost_months = extract_life_costs(months)
    for m in months:
        m.pop("raw_payments", None)

    split_items = extract_split_items(split_sheet) if split_sheet else []
    card_schedules = extract_card_schedules(split_sheet) if split_sheet else []

    # Normalize the requested target names. R is the workbook abbreviation for Rakuten.
    normalized_cards = []
    aliases = {"イオン": "イオンカード", "R": "楽天カード"}
    for item in card_schedules:
        name = aliases.get(item["name"], item["name"])
        if name in {"イオンカード", "セゾンカード", "PayPayカード", "UFJ Nicos", "楽天カード"}:
            normalized_cards.append({"name": name, "引落": item["引落"], "締め日": item["締め日"]})
    existing = {x["name"] for x in normalized_cards}
    for name in ["イオンカード", "セゾンカード", "PayPayカード", "UFJ Nicos", "楽天カード"]:
        if name not in existing:
            normalized_cards.append({"name": name, "引落": "記載なし", "締め日": "記載なし"})

    services = []
    for name in ["モビット", "千葉銀"]:
        services.append({"name": name, "引落": "記載なし", "締め日": "記載なし"})

    result = {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "source": book.name,
        "rules": {
            "incomeColumns": ["A", "B"],
            "incomeExcludedFrom": "2026-07",
            "excludedColumnsFrom2026-07": ["A", "B", "C"],
            "balanceCell": "B7",
            "paymentDetailColumn": "E",
            "paymentAmountColumn": "G",
            "installmentColumns": ["H", "I"],
            "completedRows": "分割分シートのグレーアウト行は除外",
        },
        "cardSchedules": normalized_cards,
        "serviceSchedules": services,
        "months": months,
        "lifeCostMonths": life_cost_months,
        "splitItems": split_items,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Generated {OUT} from {book.name}; monthly sheets={len(months)}; split items={len(split_items)}")


if __name__ == "__main__":
    main()

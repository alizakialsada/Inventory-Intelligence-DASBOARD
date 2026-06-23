```python
import json
import math
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA_FILE = ROOT / "dashboard-data.js"
E200_FILE = ROOT / "incoming" / "E200_latest.xls"
E100_FILE = ROOT / "incoming" / "E100_latest.xls"
INFO_FILE = ROOT / "incoming" / "update-info.json"


def clean_code(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(int(round(value)))
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def norm(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def find_header(raw):
    for i in range(min(40, len(raw))):
        row_text = " | ".join(norm(x) for x in raw.iloc[i].tolist())
        if "generic" in row_text and "available" in row_text:
            return i
    return None


def read_excel_like(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Missing required file: {path}")

    # Try Excel engines first.
    try:
        xls = pd.ExcelFile(path)
        for sheet in xls.sheet_names:
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
            header_row = find_header(raw)
            if header_row is not None:
                df = pd.read_excel(path, sheet_name=sheet, header=header_row)
                return df
    except Exception:
        pass

    # Some .xls reports are actually HTML tables with .xls extension.
    try:
        tables = pd.read_html(path)
        for raw in tables:
            header_row = find_header(raw)
            if header_row is not None:
                cols = raw.iloc[header_row].tolist()
                df = raw.iloc[header_row + 1:].copy()
                df.columns = cols
                return df
    except Exception as exc:
        raise RuntimeError(f"Could not read {path.name}: {exc}")

    raise RuntimeError(f"Could not find usable inventory table in {path.name}")


def read_stock(path: Path):
    """
    Returns:
      stock: dict Generic Code -> Available Qty
      names: dict Generic Code -> Generic Description
    """
    df = read_excel_like(path)
    columns = list(df.columns)
    normalized = [norm(c) for c in columns]

    generic_col = None
    qty_col = None
    desc_col = None

    for col, n in zip(columns, normalized):
        if generic_col is None and "generic" in n and "description" not in n:
            generic_col = col

        if qty_col is None and (
            "available qty" in n
            or "available quantity" in n
            or n == "qty"
            or n.endswith(" qty")
        ):
            qty_col = col

        if desc_col is None and (
            "generic description" in n
            or "material desc" in n
            or n == "description"
            or n.endswith(" description")
        ):
            desc_col = col

    if generic_col is None or qty_col is None:
        raise RuntimeError(f"Required columns not found in {path.name}. Columns: {columns}")

    stock = {}
    names = {}

    for _, row in df.iterrows():
        code = clean_code(row.get(generic_col))
        qty = pd.to_numeric(row.get(qty_col), errors="coerce")

        if not code:
            continue

        if desc_col is not None:
            desc = str(row.get(desc_col) or "").strip()
            if desc and desc.lower() != "nan":
                names.setdefault(code, desc)

        if pd.isna(qty) or float(qty) <= 0:
            continue

        stock[code] = stock.get(code, 0.0) + float(qty)

    return stock, names


def parse_data_js(text: str) -> dict:
    text = text.strip()
    text = re.sub(r"^window\.DASHBOARD_DB\s*=\s*", "", text)
    text = re.sub(r";\s*$", "", text)
    return json.loads(text)


def source(lc, mosool):
    if lc > 0 and mosool > 0:
        return "LC + Mosool"
    if lc > 0:
        return "LC only"
    if mosool > 0:
        return "Mosool only"
    return "Not available"


def status(total, need):
    if total <= 0:
        return "Not Available"
    if need > 0 and total < need:
        return "Below Recommended Need"
    return "Covered"


def priority_label(score):
    if score >= 85:
        return "Urgent"
    if score >= 65:
        return "High"
    if score >= 35:
        return "Monitor"
    return "Stable"


def round2(x):
    return round(float(x or 0), 2)


def percentile(values, p):
    values = sorted([float(v) for v in values if float(v) > 0])
    if not values:
        return 1.0
    idx = (len(values) - 1) * p
    lo = math.floor(idx)
    hi = math.ceil(idx)
    if lo == hi:
        return values[lo]
    return values[lo] + (values[hi] - values[lo]) * (idx - lo)


def add_inventory_only_items(db, e200, e100, e200_names, e100_names):
    """
    Adds inventory items that exist in E200/E100 but do not exist in the dashboard index.
    These items have stock but no department demand history.
    """
    existing_codes = set(str(item.get("generic", "")).strip() for item in db.get("items", []))
    all_stock_codes = set(e200.keys()) | set(e100.keys())

    added = 0

    for code in sorted(all_stock_codes):
        if not code or code in existing_codes:
            continue

        lc = float(e200.get(code, 0.0) or 0.0)
        mosool = float(e100.get(code, 0.0) or 0.0)
        total = lc + mosool

        if total <= 0:
            continue

        name = (
            e200_names.get(code)
            or e100_names.get(code)
            or "Inventory item"
        )

        item = {
            "generic": code,
            "name": str(name).strip(),
            "total_requested": 0,
            "dept_count": 0,
            "recommended_weekly_need": 0,
            "peak_weekly_sum": 0,
            "lc_qty": int(round(lc)),
            "mosool_qty": int(round(mosool)),
            "total_qty": int(round(total)),
            "coverage_weeks": 0,
            "coverage_days": 0,
            "gap_qty_total": 0,
            "source": source(lc, mosool),
            "status": "Covered",
            "priority_score": 0,
            "priority_label": "Stable",
            "patterns": "Inventory only / No department demand"
        }

        db["items"].append(item)
        existing_codes.add(code)
        added += 1

    db["overall"]["inventory_only_added"] = added


def update_db(db, e200, e100, e200_names, e100_names, latest_text):
    q90 = percentile([x.get("recommended_weekly_need", 0) for x in db["items"]], 0.90) or 1

    # Update department-level items using latest E200/E100 stock.
    for row in db["deptItems"]:
        code = str(row.get("generic", ""))
        lc = e200.get(code, 0.0)
        mosool = e100.get(code, 0.0)
        total = lc + mosool
        need = float(row.get("recommended_weekly_need", 0) or 0)

        row["lc_qty"] = int(round(lc))
        row["mosool_qty"] = int(round(mosool))
        row["total_qty"] = int(round(total))
        row["coverage_days"] = round2((total / need) * 7) if need > 0 else 0
        row["gap_qty"] = max(0, round2(need - total))
        row["source"] = source(lc, mosool)
        row["status"] = status(total, need)

    # Update existing index items using latest E200/E100 stock.
    for item in db["items"]:
        code = str(item.get("generic", ""))
        lc = e200.get(code, 0.0)
        mosool = e100.get(code, 0.0)
        total = lc + mosool
        need = float(item.get("recommended_weekly_need", 0) or 0)

        item["lc_qty"] = int(round(lc))
        item["mosool_qty"] = int(round(mosool))
        item["total_qty"] = int(round(total))
        item["coverage_weeks"] = round2(total / need) if need > 0 else 0
        item["coverage_days"] = round2((total / need) * 7) if need > 0 else 0
        item["gap_qty_total"] = max(0, round2(need - total))
        item["source"] = source(lc, mosool)
        item["status"] = status(total, need)

        base = 70 if item["status"] == "Not Available" else 45 if item["status"] == "Below Recommended Need" else 10
        dept_boost = min(20, int(item.get("dept_count", 0) or 0) * 2)
        demand_boost = min(10, (need / q90) * 10)
        score = int(round(base + dept_boost + demand_boost))
        item["priority_score"] = score
        item["priority_label"] = priority_label(score)

    # Add new inventory-only items that are not in the original index.
    add_inventory_only_items(db, e200, e100, e200_names, e100_names)

    # Sort items after adding new inventory-only items.
    db["items"].sort(key=lambda x: str(x.get("generic", "")))

    # Recalculate department summary.
    for dep in db["deptSummary"]:
        rows = [r for r in db["deptItems"] if r.get("department") == dep.get("department")]
        total = len(rows)
        covered = sum(1 for r in rows if r.get("status") == "Covered")
        below = sum(1 for r in rows if r.get("status") == "Below Recommended Need")
        na = sum(1 for r in rows if r.get("status") == "Not Available")
        gap = sum(float(r.get("gap_qty", 0) or 0) for r in rows)

        dep["unique_items"] = total
        dep["covered_items"] = covered
        dep["below_need"] = below
        dep["not_available"] = na
        dep["gap_qty"] = round2(gap)
        dep["readiness"] = round2((covered / total) * 100) if total else 0

    db["deptSummary"].sort(key=lambda d: (d.get("readiness", 0), -d.get("not_available", 0)))

    requested = len(db["items"])
    covered = sum(1 for x in db["items"] if x.get("status") == "Covered")
    below = sum(1 for x in db["items"] if x.get("status") == "Below Recommended Need")
    na = sum(1 for x in db["items"] if x.get("status") == "Not Available")

    db["overall"]["requested_items"] = requested
    db["overall"]["covered"] = covered
    db["overall"]["below"] = below
    db["overall"]["not_available"] = na
    db["overall"]["critical"] = below + na
    db["overall"]["readiness"] = round2((covered / requested) * 100) if requested else 0

    db["overall"]["lc_stock"] = int(round(sum(float(x.get("lc_qty", 0) or 0) for x in db["items"])))
    db["overall"]["mosool_stock"] = int(round(sum(float(x.get("mosool_qty", 0) or 0) for x in db["items"])))
    db["overall"]["total_stock"] = int(round(sum(float(x.get("total_qty", 0) or 0) for x in db["items"])))

    db["overall"]["last_inventory_update"] = latest_text


def main():
    e200, e200_names = read_stock(E200_FILE)
    e100, e100_names = read_stock(E100_FILE)

    latest_text = "Latest email update"
    if INFO_FILE.exists():
        try:
            info = json.loads(INFO_FILE.read_text(encoding="utf-8"))
            latest_text = info.get("latestDateText") or latest_text
        except Exception:
            pass

    db = parse_data_js(DATA_FILE.read_text(encoding="utf-8"))

    before_items = len(db.get("items", []))

    update_db(db, e200, e100, e200_names, e100_names, latest_text)

    after_items = len(db.get("items", []))
    added_items = after_items - before_items

    DATA_FILE.write_text(
        "window.DASHBOARD_DB = "
        + json.dumps(db, ensure_ascii=False, separators=(",", ":"))
        + ";\n",
        encoding="utf-8"
    )

    print(
        f"Updated dashboard-data.js. "
        f"E200 codes={len(e200)}, "
        f"E100 codes={len(e100)}, "
        f"items before={before_items}, "
        f"items after={after_items}, "
        f"inventory-only added={added_items}"
    )


if __name__ == "__main__":
    main()
```

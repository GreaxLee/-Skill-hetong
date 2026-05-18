#!/usr/bin/env python3
"""
generate_contract.py — Automatically generate purchase contracts from an order plan.

Usage:
    python generate_contract.py 2026-06-01
"""

import sys
import os
import re
import shutil
import zipfile
import datetime
from copy import copy
from typing import Optional

import openpyxl
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

ORDER_PLAN_FILE = "/Users/shawlee/Desktop/订货计划共享表.xlsx"
PRICE_LIST_FILE = "/Users/shawlee/Desktop/产品价格表.xlsx"

# Folders to scan for existing contracts (SKU detail lookup)
DETAIL_SCAN_DIRS = [
    "/Users/shawlee/好消息/工厂/惠州丽瑞/合同/",
    "/Users/shawlee/好消息/工厂/东莞永铭/合同/",
    "/Users/shawlee/好消息/工厂/青岛翊霖/合同/",
]

EXCEL_DATE_ORIGIN = datetime.date(1899, 12, 30)

FACTORY_CONFIG = {
    "惠州市丽瑞家居有限公司": {
        "contract_prefix": "01HZLR",
        "base_folder": "/Users/shawlee/好消息/工厂/惠州丽瑞/合同",
        "file_prefix": "丽瑞合同",
        "seller_name": "惠州市丽瑞家居有限公司",
        "seller_contact": "刘小琼",
        "seller_phone": "13725559521",
        "bank_name": "招商银行惠州惠阳支行",
        "bank_account": "752901874210268",
        "bank_holder": "惠州市丽瑞家居有限公司",
        "payment_terms": "货款月结，次月25号",
        "buyer_contact": "王语嫣",
        "buyer_phone": "15720616686",
        "price_col": "含税价",
    },
    "东莞市永铭五金制品有限公司": {
        "contract_prefix": "01DGYM",
        "base_folder": "/Users/shawlee/好消息/工厂/东莞永铭/合同",
        "file_prefix": "永铭合同",
        "seller_name": "东莞市永铭五金制品有限公司",
        "seller_contact": "林建民",
        "seller_phone": "13902311630",
        "bank_name": "中国银行东莞清溪支行",
        "bank_account": "719857882643",
        "bank_holder": "东莞市永铭五金制品有限公司",
        "payment_terms": "货款月结，次月20号",
        "buyer_contact": "王语嫣",
        "buyer_phone": "15720616686",
        "price_col": "含税价",
    },
    "青岛翊霖源工艺品有限公司": {
        "contract_prefix": "01QYLY",
        "base_folder": "/Users/shawlee/好消息/工厂/青岛翊霖/合同",
        "file_prefix": "翊霖源合同",
        "seller_name": "青岛翊霖源工艺品有限公司",
        "seller_contact": "王洪志",
        "seller_phone": "15898873161",
        "bank_name": "青岛银行",
        "bank_account": "6232120100006539858",
        "bank_holder": "王守军",
        "payment_terms": "出货后30天内支付剩余货款",
        "buyer_contact": "郭重阳",
        "buyer_phone": "17203878583",
        "price_col": "含税价",
    },
}


def get_year_folder(factory_name: str, year: int) -> str:
    """Return the contract folder for the given year, creating it if needed."""
    cfg = FACTORY_CONFIG[factory_name]
    year_short = str(year)[-2:]
    folder = os.path.join(cfg["base_folder"], f"{year_short}年")
    os.makedirs(folder, exist_ok=True)
    return folder


def get_template_file(factory_name: str, year: int) -> Optional[str]:
    """Find the latest contract in the current year's folder to use as template.
    Falls back to the previous year if no contracts exist yet for this year."""
    cfg = FACTORY_CONFIG[factory_name]
    prefix = cfg["file_prefix"]
    pattern = re.compile(rf"^{re.escape(prefix)}\d+\.xlsx$")

    for search_year in (year, year - 1):
        year_short = str(search_year)[-2:]
        folder = os.path.join(cfg["base_folder"], f"{year_short}年")
        if not os.path.isdir(folder):
            continue
        candidates = [f for f in os.listdir(folder) if pattern.match(f)]
        if candidates:
            latest = sorted(candidates)[-1]
            return os.path.join(folder, latest)

    return None


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------


def excel_serial_to_date(serial) -> Optional[datetime.date]:
    """Convert an Excel serial date number to a Python date."""
    if serial is None:
        return None
    try:
        n = int(serial)
    except (TypeError, ValueError):
        return None
    # Excel incorrectly treats 1900 as a leap year; serials <= 59 are pre-March-1900
    if n <= 0:
        return None
    return EXCEL_DATE_ORIGIN + datetime.timedelta(days=n)


def normalize_date(value) -> Optional[datetime.date]:
    """Accept either a serial int or a datetime/date object."""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.date() if isinstance(value, datetime.datetime) else value
    if isinstance(value, (int, float)):
        return excel_serial_to_date(value)
    return None


# ---------------------------------------------------------------------------
# Price list loading
# ---------------------------------------------------------------------------


def load_price_list(path: str) -> dict:
    """
    Returns a dict: sku (str) -> {"出厂价": float|None, "含税价": float|None}
    Price list SKUs may be combined like "N03SR018BB/BR /WW"; we expand them.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    prices: dict[str, dict] = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        sku_raw, factory, factory_price, tax_price = (row[0], row[1], row[2], row[3])
        if not sku_raw:
            continue
        sku_raw = str(sku_raw).strip()

        # Expand combined SKUs like "N03SR018BB/BR /WW"
        variants = _expand_sku_variants(sku_raw)
        for sku in variants:
            prices[sku] = {
                "出厂价": float(factory_price) if factory_price is not None else None,
                "含税价": float(tax_price) if tax_price is not None else None,
                "工厂": str(factory).strip() if factory else "",
            }

    return prices


def _expand_sku_variants(sku_raw: str) -> list[str]:
    """
    'N03SR018BB/BR /WW' -> ['N03SR018BB', 'N03SR018BR', 'N03SR018WW']
    'N02LH003NB'        -> ['N02LH003NB']
    """
    cleaned = sku_raw.replace(" ", "")
    parts = cleaned.split("/")
    if len(parts) == 1:
        return [parts[0]]
    first = parts[0]
    # suffix length is 2 chars for color codes
    prefix = first[:-2]
    variants = [first] + [prefix + s for s in parts[1:] if s]
    return variants


# ---------------------------------------------------------------------------
# SKU detail lookup (scan existing contracts)
# ---------------------------------------------------------------------------


def build_sku_detail_lookup(scan_dirs: list[str]) -> dict:
    """
    Scan all .xlsx files under scan_dirs for sheets named '合同标的'.
    Returns dict: sku -> {"产品名称": str, "参数": str}

    Two header layouts are handled:
      New (26年): col A=SKU, B=产品名称, C=图片, D=参数
      Old (25年): col A=SKU, B=FNSKU, C=产品名称, D=图片, E=参数
    """
    lookup: dict[str, dict] = {}

    for scan_dir in scan_dirs:
        for root, _dirs, files in os.walk(scan_dir):
            for fname in files:
                if not fname.endswith(".xlsx") or fname.startswith("~"):
                    continue
                fpath = os.path.join(root, fname)
                try:
                    wb = openpyxl.load_workbook(fpath, data_only=True)
                except Exception as e:
                    print(f"  [警告] 无法读取 {fpath}: {e}")
                    continue
                if "合同标的" not in wb.sheetnames:
                    continue
                ws = wb["合同标的"]
                _extract_sku_details(ws, lookup, fpath)

    return lookup


def _extract_sku_details(ws, lookup: dict, fpath: str):
    """Parse one 合同标的 worksheet and populate lookup."""
    # Detect header layout by inspecting row 5
    h = {ws.cell(5, c).value: c for c in range(1, 12) if ws.cell(5, c).value}

    if "FNSKU" in h:
        # Old layout: A=SKU, B=FNSKU, C=产品名称, D=图片, E=参数
        col_name = 3
        col_param = 5
    else:
        # New layout: A=SKU, B=产品名称, C=图片, D=参数
        col_name = 2
        col_param = 4

    for row in ws.iter_rows(min_row=6, values_only=True):
        sku = row[0]
        if not sku:
            continue
        sku = str(sku).strip()
        if not sku or not sku[0].isalnum():
            continue

        name_val = row[col_name - 1]
        param_val = row[col_param - 1]
        name = str(name_val).strip() if name_val else ""
        param = str(param_val).strip() if param_val else ""

        # Prefer the most recent entry (overwrite is fine as we scan newest last)
        if sku not in lookup or (name and not lookup[sku]["产品名称"]):
            lookup[sku] = {"产品名称": name, "参数": param}


# ---------------------------------------------------------------------------
# Contract number generation
# ---------------------------------------------------------------------------


def next_contract_number(factory_name: str, year: int) -> tuple[str, str, int]:
    """
    Scan the current year's contract folder for existing files.
    Returns (contract_number, filename, seq_int).
    """
    cfg = FACTORY_CONFIG[factory_name]
    folder = get_year_folder(factory_name, year)
    prefix = cfg["file_prefix"]
    contract_prefix = cfg["contract_prefix"]
    year_short = str(year)[-2:]

    max_seq = 0
    pattern = re.compile(rf"^{re.escape(prefix)}(\d+)\.xlsx$")
    for fname in os.listdir(folder):
        m = pattern.match(fname)
        if m:
            num_str = m.group(1)
            try:
                seq = int(num_str[-3:])
                max_seq = max(max_seq, seq)
            except ValueError:
                pass

    next_seq = max_seq + 1
    contract_number = f"{contract_prefix}{year}{next_seq:03d}"
    filename = f"{prefix}{year_short}{next_seq:03d}.xlsx"
    return contract_number, filename, next_seq


# ---------------------------------------------------------------------------
# Order plan parsing
# ---------------------------------------------------------------------------


def parse_order_plan(path: str, target_date: datetime.date) -> dict:
    """
    Returns orders grouped by factory:
      { factory_name: [ {"sku": str, "qty": int, "chinese_name": str, "source_sheet": str} ] }
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    all_orders: dict[str, list] = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_orders = _parse_sheet(ws, sheet_name, target_date)
        for factory, items in sheet_orders.items():
            all_orders.setdefault(factory, []).extend(items)

    return all_orders


def _parse_sheet(ws, sheet_name: str, target_date: datetime.date) -> dict:
    """Dispatch to the correct parser based on sheet name."""
    if sheet_name == "脏衣篮订货":
        return _parse_zang_yi_lan(ws, sheet_name, target_date)
    elif sheet_name in ("铁网鞋架", "铁线"):
        return _parse_tie_wang_or_xian(ws, sheet_name, target_date)
    elif sheet_name == "Larzonic":
        return _parse_larzonic(ws, sheet_name, target_date)
    else:
        print(f"  [警告] 未知sheet: {sheet_name}，跳过")
        return {}


def _parse_zang_yi_lan(ws, sheet_name: str, target_date: datetime.date) -> dict:
    """
    脏衣篮订货:
      header row=8, shipment date row=4, data starts row=9
      date cols start col=13 (odd cols=date, even cols=箱数/ignored)
      SKU=col B (2), factory=col C (3)
    """
    orders: dict[str, list] = {}

    # Collect date columns: col 13, 15, 17, … up to max_col
    date_cols = []
    for c in range(13, ws.max_column + 1, 2):  # odd cols = date
        raw = ws.cell(4, c).value
        d = normalize_date(raw)
        if d is not None:
            date_cols.append((c, d))

    # Find columns matching target_date
    target_cols = [c for c, d in date_cols if d == target_date]
    if not target_cols:
        return {}

    for row in ws.iter_rows(min_row=9, values_only=True):
        sku = row[1]  # col B (0-indexed: 1)
        factory = row[2]  # col C
        if not sku or not factory:
            continue
        sku = str(sku).strip()
        factory = str(factory).strip()
        if not sku or not sku[0].isalnum():
            continue

        for c in target_cols:
            qty_raw = row[c - 1]  # 0-indexed
            if qty_raw is None:
                continue
            try:
                qty = int(qty_raw)
            except (TypeError, ValueError):
                continue
            if qty <= 0:
                continue
            orders.setdefault(factory, []).append({
                "sku": sku,
                "qty": qty,
                "chinese_name": "",
                "source_sheet": sheet_name,
            })

    return _merge_duplicate_skus(orders)


def _parse_tie_wang_or_xian(ws, sheet_name: str, target_date: datetime.date) -> dict:
    """
    铁网鞋架 / 铁线:
      header row=9, shipment date row=5, data starts row=10
      date cols start col=10 (even cols=date, odd cols=箱数/ignored)
      SKU=col A (1), factory=col B (2), 中文名=col C (3)
    """
    orders: dict[str, list] = {}

    # Collect date columns: col 10, 12, 14, … (even cols = date)
    date_cols = []
    for c in range(10, ws.max_column + 1, 2):
        raw = ws.cell(5, c).value
        d = normalize_date(raw)
        if d is not None:
            date_cols.append((c, d))

    target_cols = [c for c, d in date_cols if d == target_date]
    if not target_cols:
        return {}

    for row in ws.iter_rows(min_row=10, values_only=True):
        sku = row[0]   # col A
        factory = row[1]  # col B
        chinese_name = row[2] or ""  # col C
        if not sku or not factory:
            continue
        sku = str(sku).strip()
        factory = str(factory).strip()
        chinese_name = str(chinese_name).strip()
        if not sku or not sku[0].isalnum():
            continue

        for c in target_cols:
            qty_raw = row[c - 1]
            if qty_raw is None:
                continue
            try:
                qty = int(qty_raw)
            except (TypeError, ValueError):
                continue
            if qty <= 0:
                continue
            orders.setdefault(factory, []).append({
                "sku": sku,
                "qty": qty,
                "chinese_name": chinese_name,
                "source_sheet": sheet_name,
            })

    return _merge_duplicate_skus(orders)


def _parse_larzonic(ws, sheet_name: str, target_date: datetime.date) -> dict:
    """
    Larzonic:
      header row=6, shipment date row=3, data starts row=7
      date cols start col=11 (all cols=数量, no 箱数)
      SKU=col A (1), factory=col B (2), 中文名=col C (3)
      Mark items with larzonic=True so we know to use 出厂价.
    """
    orders: dict[str, list] = {}

    # Collect date columns: col 11, 12, 13, … all are qty columns
    date_cols = []
    for c in range(11, ws.max_column + 1):
        raw = ws.cell(3, c).value
        d = normalize_date(raw)
        if d is not None:
            date_cols.append((c, d))

    target_cols = [c for c, d in date_cols if d == target_date]
    if not target_cols:
        return {}

    for row in ws.iter_rows(min_row=7, values_only=True):
        sku = row[0]
        factory = row[1]
        chinese_name = row[2] or ""
        if not sku or not factory:
            continue
        sku = str(sku).strip()
        factory = str(factory).strip()
        chinese_name = str(chinese_name).strip()
        if not sku or not sku[0].isalnum():
            continue

        for c in target_cols:
            qty_raw = row[c - 1]
            if qty_raw is None:
                continue
            try:
                qty = int(qty_raw)
            except (TypeError, ValueError):
                continue
            if qty <= 0:
                continue
            orders.setdefault(factory, []).append({
                "sku": sku,
                "qty": qty,
                "chinese_name": chinese_name,
                "source_sheet": sheet_name,
                "larzonic": True,  # use 出厂价
            })

    return _merge_duplicate_skus(orders)


def _merge_duplicate_skus(orders: dict) -> dict:
    """Merge entries with the same SKU within each factory (sum quantities)."""
    merged: dict[str, list] = {}
    for factory, items in orders.items():
        seen: dict[str, dict] = {}
        for item in items:
            sku = item["sku"]
            if sku in seen:
                seen[sku]["qty"] += item["qty"]
            else:
                seen[sku] = dict(item)
        merged[factory] = list(seen.values())
    return merged


# ---------------------------------------------------------------------------
# Contract generation
# ---------------------------------------------------------------------------


def generate_contract(
    factory_name: str,
    orders: list,
    target_date: datetime.date,
    price_lookup: dict,
    sku_detail_lookup: dict,
    today: datetime.date,
) -> Optional[str]:
    """
    Generate one contract Excel file for factory_name.
    Returns the output filename on success, None if no valid SKUs.
    """
    cfg = FACTORY_CONFIG[factory_name]
    year = today.year
    contract_number, filename, seq = next_contract_number(factory_name, year)
    folder = get_year_folder(factory_name, year)
    out_path = os.path.join(folder, filename)

    template_path = get_template_file(factory_name, year)
    if template_path is None:
        print(f"  [错误] 找不到 {factory_name} 的合同模板，跳过")
        return None

    # Sort orders by SKU
    orders_sorted = sorted(orders, key=lambda x: x["sku"])

    # Filter & enrich orders with prices
    enriched = []
    for item in orders_sorted:
        sku = item["sku"]
        is_larzonic = item.get("larzonic", False)

        price_info = price_lookup.get(sku)
        if price_info is None:
            print(f"  [警告] 价格表中未找到SKU: {sku}，跳过")
            continue

        if is_larzonic:
            price = price_info.get("出厂价")
        else:
            price_col_key = cfg["price_col"]  # "含税价" or "出厂价"
            price = price_info.get(price_col_key)

        if price is None:
            print(f"  [警告] SKU {sku} 的价格为空，跳过")
            continue

        detail = sku_detail_lookup.get(sku)
        if detail is None:
            print(f"  [警告] SKU详情未找到: {sku}，将使用中文名")
            product_name = item.get("chinese_name", "")
            param = ""
        else:
            product_name = detail["产品名称"] or item.get("chinese_name", "")
            param = detail["参数"]

        enriched.append({
            "sku": sku,
            "product_name": product_name,
            "param": param,
            "qty": item["qty"],
            "price": price,
        })

    if not enriched:
        print(f"  [跳过] {factory_name}: 没有有效SKU")
        return None

    # Copy template
    shutil.copy2(template_path, out_path)

    wb = openpyxl.load_workbook(out_path)

    # ---- 采购单 sheet ----
    ws_purchase = wb["采购单"]
    _fill_purchase_sheet(ws_purchase, contract_number, today)

    # ---- 合同标的 sheet ----
    ws_contract = wb["合同标的"]
    _fill_contract_sheet(ws_contract, contract_number, target_date, enriched)

    wb.save(out_path)

    # Copy drawings and media from template (openpyxl drops embedded images on save)
    _copy_drawings_from_template(template_path, out_path)

    total_qty = sum(item["qty"] for item in enriched)
    total_amount = sum(item["qty"] * item["price"] for item in enriched)
    print(
        f"生成合同：{filename}  ({factory_name}, {len(enriched)}个SKU, "
        f"合计{total_qty}套, ¥{total_amount:,.0f})"
    )
    return filename


def _fill_purchase_sheet(ws, contract_number: str, today: datetime.date):
    """Update 采购单: contract number (B2) and signing date in both buyer and seller cells."""
    # Contract number is in B2
    ws.cell(2, 2).value = contract_number

    # Both buyer (A22) and seller (F22) cells contain "签订日期：YYYY.M.D"
    new_date_str = f"{today.year}.{today.month}.{today.day}"
    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row_idx, col_idx)
            if cell.value and "签订日期" in str(cell.value):
                cell.value = re.sub(
                    r"(签订日期[：:])[\d./]+",
                    rf"\g<1>{new_date_str}",
                    str(cell.value),
                )


def _fill_contract_sheet(
    ws,
    contract_number: str,
    target_date: datetime.date,
    enriched: list,
):
    """Update 合同标的: contract number, delivery date, and data rows."""
    # Row 1: 订单号 / contract number
    ws.cell(1, 2).value = contract_number

    # Row 2: 交货期 / delivery date — write as datetime so Excel formats it correctly
    ws.cell(2, 2).value = datetime.datetime(target_date.year, target_date.month, target_date.day)

    # Unmerge all merged ranges in the data area (row 6+) to allow free writing
    data_area_merges = [mr for mr in list(ws.merged_cells.ranges) if mr.min_row >= 6]
    for mr in data_area_merges:
        ws.unmerge_cells(str(mr))

    # Clear existing data rows (row 6 onward)
    _clear_data_rows(ws, start_row=6)

    # Write data rows starting at row 6
    data_start_row = 6
    for i, item in enumerate(enriched):
        r = data_start_row + i
        ws.cell(r, 1).value = item["sku"]
        ws.cell(r, 2).value = item["product_name"]
        # col C (3) = image, leave empty
        ws.cell(r, 4).value = item["param"]
        # col E (5): re-merge D:E for the 参数 cell (matches template style)
        ws.merge_cells(f"D{r}:E{r}")
        ws.cell(r, 6).value = item["qty"]
        ws.cell(r, 7).value = item["price"]
        ws.cell(r, 8).value = f"=F{r}*G{r}"

    # Total row
    last_data_row = data_start_row + len(enriched) - 1
    total_row = last_data_row + 1
    ws.cell(total_row, 1).value = "合计"
    ws.cell(total_row, 6).value = f"=SUM(F{data_start_row}:F{last_data_row})"
    ws.cell(total_row, 8).value = f"=SUM(H{data_start_row}:H{last_data_row})"


def _clear_data_rows(ws, start_row: int):
    """Clear all cell values from start_row to the end of data (skip merged cells)."""
    from openpyxl.cell.cell import MergedCell
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None


def _copy_drawings_from_template(template_path: str, out_path: str):
    """
    openpyxl drops xl/drawings/ and xl/media/ when saving, and also removes the
    <drawing r:id="rId1"/> reference from the worksheet XML.
    This function restores all three pieces from the template.
    """
    DRAWING_PREFIXES = ("xl/drawings/", "xl/media/")

    with zipfile.ZipFile(template_path, "r") as tmpl_zip:
        # Collect drawing/media files to copy
        drawing_names = [n for n in tmpl_zip.namelist()
                         if any(n.startswith(p) for p in DRAWING_PREFIXES)]
        if not drawing_names:
            return
        drawing_data = {name: tmpl_zip.read(name) for name in drawing_names}

        # Find which sheet XML in the template contains a <drawing ...> element
        # and capture the drawing reference XML fragment (e.g. <drawing r:id="rId1"/>)
        sheet_drawing_ref: dict[str, str] = {}  # sheet filename -> drawing XML fragment
        for name in tmpl_zip.namelist():
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                content = tmpl_zip.read(name).decode("utf-8")
                m = re.search(r'<drawing[^/]*/>', content)
                if m:
                    sheet_drawing_ref[os.path.basename(name)] = m.group(0)

        # Worksheet .rels files (needed so the sheet can resolve the drawing rId)
        sheet_rels = {name: tmpl_zip.read(name) for name in tmpl_zip.namelist()
                      if name.startswith("xl/worksheets/_rels/")}

    # Rewrite the output zip
    tmp_path = out_path + ".tmp"
    with zipfile.ZipFile(out_path, "r") as out_zip, \
         zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as new_zip:

        skip = set(drawing_data) | set(sheet_rels)
        for item in out_zip.infolist():
            if item.filename in skip:
                continue
            content = out_zip.read(item.filename)
            # Inject <drawing .../> into any sheet XML that needs it
            fname = os.path.basename(item.filename)
            if fname in sheet_drawing_ref and item.filename.startswith("xl/worksheets/"):
                text = content.decode("utf-8")
                if "<drawing" not in text:
                    drawing_tag = sheet_drawing_ref[fname]
                    # Ensure the r: namespace is declared on the root element
                    r_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                    if r_ns not in text:
                        text = text.replace("<worksheet ", f"<worksheet {r_ns} ", 1)
                    text = text.replace("</worksheet>", f"{drawing_tag}</worksheet>")
                    content = text.encode("utf-8")
            new_zip.writestr(item, content)

        for name, data in drawing_data.items():
            new_zip.writestr(name, data)
        for name, data in sheet_rels.items():
            new_zip.writestr(name, data)

    os.replace(tmp_path, out_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    if len(sys.argv) < 2:
        print("用法: python generate_contract.py <交货日期>")
        print("示例: python generate_contract.py 2026-06-01")
        sys.exit(1)

    date_str = sys.argv[1]
    try:
        target_date = datetime.date.fromisoformat(date_str)
    except ValueError:
        print(f"日期格式错误: {date_str}，请使用 YYYY-MM-DD 格式")
        sys.exit(1)

    today = datetime.date.today()
    print(f"目标交货日期: {target_date}  |  今日: {today}")
    print()

    # 1. Load price list
    print("正在加载价格表…")
    price_lookup = load_price_list(PRICE_LIST_FILE)
    print(f"  加载了 {len(price_lookup)} 个SKU的价格")

    # 2. Build SKU detail lookup from existing contracts
    print("正在扫描合同文件以建立SKU详情库…")
    sku_detail_lookup = build_sku_detail_lookup(DETAIL_SCAN_DIRS)
    print(f"  扫描到 {len(sku_detail_lookup)} 个SKU的详情")

    # 3. Parse order plan
    print("正在解析订货计划…")
    orders_by_factory = parse_order_plan(ORDER_PLAN_FILE, target_date)

    if not orders_by_factory:
        print(f"\n未找到 {target_date} 的订单，退出。")
        sys.exit(0)

    total_skus = sum(len(v) for v in orders_by_factory.values())
    print(f"  找到 {len(orders_by_factory)} 个工厂的订单:")
    for factory, items in orders_by_factory.items():
        print(f"    {factory}: {len(items)} 个SKU")

    print()

    # 4. Generate one contract per factory
    generated = []
    for factory_name in sorted(orders_by_factory.keys()):
        if factory_name not in FACTORY_CONFIG:
            print(f"  [警告] 未配置工厂: {factory_name}，跳过")
            continue
        items = orders_by_factory[factory_name]
        result = generate_contract(
            factory_name=factory_name,
            orders=items,
            target_date=target_date,
            price_lookup=price_lookup,
            sku_detail_lookup=sku_detail_lookup,
            today=today,
        )
        if result:
            generated.append(result)

    if not generated:
        print("\n没有生成任何合同。")
    else:
        print(f"\n共生成 {len(generated)} 份合同。")


if __name__ == "__main__":
    main()

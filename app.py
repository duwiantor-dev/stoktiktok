import re
import copy
from io import BytesIO
from typing import Dict, Tuple, List, Optional

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


# =========================
# Helpers
# =========================
def _norm_str(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


def _norm_key(x) -> str:
    return _norm_str(x).strip().upper()


def parse_thousands_to_rp(value) -> Optional[int]:
    """
    Pricelist & Addon file: angka biasanya "tanpa 000" (contoh 9300 => 9.300.000)
    - Kalau string: ambil semua digit (23.699 -> 23699)
    - Kalau float kecil dengan desimal 3: 23.699 -> 23699
    Return: Rupiah full (x1000)
    """
    if value is None:
        return None

    # If already int-ish
    if isinstance(value, bool):
        return None

    if isinstance(value, int):
        return int(value) * 1000

    if isinstance(value, float):
        if value != value:  # NaN
            return None
        # Case like 23.699 (float) meaning 23699
        if value < 1000 and abs(value - round(value)) > 1e-9:
            thousands = int(round(value * 1000))
            return thousands * 1000
        # Case like 9300.0
        thousands = int(round(value))
        return thousands * 1000

    s = _norm_str(value)
    if not s:
        return None

    # Keep digits only (handles "23.699" or "23,699" or "Rp 23.699")
    digits = re.findall(r"\d+", s)
    if not digits:
        return None
    thousands = int("".join(digits))
    return thousands * 1000


def parse_rp(value) -> Optional[int]:
    """Parse value in mass update price column (usually already full rupiah)."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return int(value)
    if isinstance(value, float):
        if value != value:
            return None
        return int(round(value))

    s = _norm_str(value)
    if not s:
        return None
    digits = re.findall(r"\d+", s)
    if not digits:
        return None
    return int("".join(digits))


def find_header_row(
    ws: Worksheet,
    must_have_any: List[List[str]],
    scan_rows: int = 50,
) -> Optional[int]:
    """
    Scan first N rows, return first row index that satisfies:
    - for each group in must_have_any: at least 1 keyword in that group appears in row (case-insensitive).
    Example:
      must_have_any = [
         ["KODEBARANG", "KODE BARANG"],
         ["M4"]
      ]
    """
    for r in range(1, min(scan_rows, ws.max_row) + 1):
        row_vals = [(_norm_str(ws.cell(r, c).value)).upper() for c in range(1, ws.max_column + 1)]
        row_text = " | ".join(row_vals)

        ok = True
        for group in must_have_any:
            if not any(k.upper() in row_text for k in group):
                ok = False
                break
        if ok:
            return r
    return None


def map_headers(ws: Worksheet, header_row: int) -> Dict[str, int]:
    """Return dict: normalized header -> column index."""
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        key = _norm_key(v)
        if key:
            headers[key] = c
    return headers


def get_first_sheet(wb: openpyxl.Workbook) -> Worksheet:
    return wb[wb.sheetnames[0]]


def split_sku_addons(sku_full: str) -> Tuple[str, List[str]]:
    """
    "BASE+PC+BA" => ("BASE", ["PC","BA"])
    empty addon => []
    """
    parts = [p for p in sku_full.split("+") if p.strip() != ""]
    if not parts:
        return "", []
    base = parts[0].strip()
    addons = [p.strip() for p in parts[1:]]
    return base, addons


def copy_row_style(src_ws: Worksheet, src_row: int, dst_ws: Worksheet, dst_row: int, max_col: int):
    """Copy cell style (font/fill/border/number_format/alignment/protection) from src_row to dst_row."""
    for c in range(1, max_col + 1):
        src_cell = src_ws.cell(src_row, c)
        dst_cell = dst_ws.cell(dst_row, c)
        dst_cell._style = copy.copy(src_cell._style)
        dst_cell.number_format = src_cell.number_format
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.protection = copy.copy(src_cell.protection)


# =========================
# Core processing
# =========================
def load_pricelist_map(pl_bytes: bytes) -> Tuple[Dict[str, int], int]:
    """
    Returns:
      pl_map: {KODEBARANG -> price_rp_from_M4}
      header_row
    """
    wb = openpyxl.load_workbook(BytesIO(pl_bytes), data_only=True)
    ws = get_first_sheet(wb)

    header_row = find_header_row(
        ws,
        must_have_any=[["KODEBARANG", "KODE BARANG"], ["M4"]],
        scan_rows=80,
    )
    if not header_row:
        raise ValueError("Header Pricelist tidak ketemu. Pastikan ada kolom KODEBARANG & M4.")

    hdr = map_headers(ws, header_row)
    # Must have KODEBARANG
    sku_col = hdr.get("KODEBARANG") or hdr.get("KODE BARANG")
    m4_col = hdr.get("M4")
    if not sku_col or not m4_col:
        raise ValueError("Kolom SKU/Price di Pricelist tidak ketemu (butuh KODEBARANG/KODE BARANG & M4).")

    pl_map: Dict[str, int] = {}

    for r in range(header_row + 1, ws.max_row + 1):
        sku = ws.cell(r, sku_col).value
        if sku is None:
            continue
        sku_key = _norm_key(sku)
        if not sku_key or sku_key == "TOTAL":
            continue
        price_raw = ws.cell(r, m4_col).value
        price_rp = parse_thousands_to_rp(price_raw)
        if price_rp is None:
            continue
        pl_map[sku_key] = price_rp

    if not pl_map:
        raise ValueError("Pricelist kebaca, tapi mapping KODEBARANG -> M4 kosong. Cek isi file Pricelist.")
    return pl_map, header_row


def load_addon_map(addon_bytes: bytes) -> Dict[str, int]:
    """
    Addon file expected to have a code column and a price column.
    Auto-detect columns by keywords:
      - code: contains "ADDON", "KODE", "VARIAN", "STANDARISASI"
      - price: contains "HARGA", "PRICE"
    Return: {ADDON_CODE -> addon_price_rp}
    """
    wb = openpyxl.load_workbook(BytesIO(addon_bytes), data_only=True)
    ws = get_first_sheet(wb)

    header_row = find_header_row(
        ws,
        must_have_any=[["HARGA", "PRICE"], ["KODE", "VARIAN", "STANDARISASI", "ADDON"]],
        scan_rows=80,
    )
    if not header_row:
        raise ValueError("Header Addon Mapping tidak ketemu. Pastikan ada kolom kode addon & harga.")

    hdr = map_headers(ws, header_row)

    # find price col
    price_col = None
    for k in ["HARGA", "PRICE"]:
        if k in hdr:
            price_col = hdr[k]
            break

    # find code col (prioritize the first column-like)
    code_col = None
    for candidate in [
        "ADDON_CODE",
        "KODE",
        "KODE ADDON",
        "STANDARISASI KODE SKU DI VARIAN",
        "STANDARISASI KODE SKU DI VARIASI",
        "KODE VARIAN",
        "KODE VARIASI",
    ]:
        if candidate in hdr:
            code_col = hdr[candidate]
            break

    if not code_col:
        # fallback: pick the leftmost header that contains "KODE" or "VARIAN" or "STANDARISASI"
        for hk, col in hdr.items():
            if any(x in hk for x in ["KODE", "VARIAN", "VARIASI", "STANDARISASI", "ADDON"]):
                code_col = col
                break

    if not code_col or not price_col:
        raise ValueError("Kolom addon_code / harga tidak ketemu di Addon Mapping.")

    addon_map: Dict[str, int] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(r, code_col).value
        if code is None:
            continue
        code_key = _norm_key(code)
        if not code_key:
            continue

        price_raw = ws.cell(r, price_col).value
        price_rp = parse_thousands_to_rp(price_raw)
        if price_rp is None:
            continue

        addon_map[code_key] = price_rp

    # It's ok if empty, user might upload none/blank
    return addon_map


def compute_final_price(
    sku_full: str,
    pricelist_map: Dict[str, int],
    addon_map: Dict[str, int],
    diskon_rp: int,
) -> Optional[int]:
    """
    Rules:
    - Base SKU must exist in pricelist_map (KODEBARANG)
    - If addons exist:
        * EVERY addon must exist in addon_map
        * else return None (means DO NOT UPDATE)
    - Final = base + sum(addons) - diskon_rp
    """
    base, addons = split_sku_addons(sku_full)
    base_key = _norm_key(base)
    if not base_key or base_key not in pricelist_map:
        return None

    base_price = pricelist_map[base_key]
    addon_sum = 0

    for a in addons:
        akey = _norm_key(a)
        if not akey:
            continue
        if akey not in addon_map:
            # STRICT: if one missing => no change
            return None
        addon_sum += addon_map[akey]

    final_price = base_price + addon_sum - int(diskon_rp)
    if final_price < 0:
        final_price = 0
    return final_price


def process_shopee_files(
    mass_files: List[Tuple[str, bytes]],
    pl_bytes: bytes,
    addon_bytes: Optional[bytes],
    diskon_rp: int,
) -> Tuple[bytes, pd.DataFrame]:
    """
    Output:
      - one XLSX bytes: template format from FIRST mass file, rows from 7 contain only updated rows
      - issues_df: rows that couldn't be updated, or file errors
    """
    pricelist_map, _ = load_pricelist_map(pl_bytes)
    addon_map: Dict[str, int] = {}
    if addon_bytes:
        addon_map = load_addon_map(addon_bytes)

    issues = []

    # Use first file as template output
    first_name, first_bytes = mass_files[0]
    out_wb = openpyxl.load_workbook(BytesIO(first_bytes))
    out_ws = get_first_sheet(out_wb)

    # Shopee template: header row 3, data starts row 7
    DATA_START_ROW = 7
    # Determine columns by fixed position (as requested): SKU col F (6), price col G (7)
    SKU_COL = 6
    PRICE_COL = 7

    # Clear existing data rows (delete from row 7 to end)
    if out_ws.max_row >= DATA_START_ROW:
        out_ws.delete_rows(DATA_START_ROW, out_ws.max_row - DATA_START_ROW + 1)

    # Capture style template from row 7 of the original first file (we reload to read style)
    tmp_wb = openpyxl.load_workbook(BytesIO(first_bytes))
    tmp_ws = get_first_sheet(tmp_wb)
    template_style_row = DATA_START_ROW  # row 7 contains first data row in template file
    max_col = tmp_ws.max_column

    write_row = DATA_START_ROW
    total_updated = 0

    for fname, fbytes in mass_files:
        try:
            wb = openpyxl.load_workbook(BytesIO(fbytes))
            ws = get_first_sheet(wb)

            # iterate rows from 7 onward until blank SKU for a while
            for r in range(DATA_START_ROW, ws.max_row + 1):
                sku_full = ws.cell(r, SKU_COL).value
                sku_full_s = _norm_str(sku_full)
                if not sku_full_s:
                    continue

                old_price = ws.cell(r, PRICE_COL).value
                old_price_rp = parse_rp(old_price)

                new_price = compute_final_price(
                    sku_full=sku_full_s,
                    pricelist_map=pricelist_map,
                    addon_map=addon_map,
                    diskon_rp=diskon_rp,
                )

                if new_price is None:
                    # Not updated: either base not in PL or addon missing -> skip
                    continue

                # Only output rows where price actually changes (or old is empty)
                if old_price_rp is not None and int(old_price_rp) == int(new_price):
                    continue

                # Write a styled row into output template
                out_ws.insert_rows(write_row, 1)
                copy_row_style(tmp_ws, template_style_row, out_ws, write_row, max_col)

                # Copy entire row values from source file to output row
                for c in range(1, max_col + 1):
                    out_ws.cell(write_row, c).value = ws.cell(r, c).value

                # Overwrite price cell only
                out_ws.cell(write_row, PRICE_COL).value = int(new_price)

                total_updated += 1
                write_row += 1

        except Exception as e:
            issues.append(
                {"file": fname, "row": "", "sku_full": "", "base_sku": "", "reason": f"Gagal proses file: {e}"}
            )

    if total_updated == 0:
        issues.append(
            {
                "file": "",
                "row": "",
                "sku_full": "",
                "base_sku": "",
                "reason": "Tidak ada baris yang berubah (semua cocok / tidak ada yang bisa diupdate).",
            }
        )

    # Save output
    out_buf = BytesIO()
    out_wb.save(out_buf)
    out_bytes = out_buf.getvalue()

    issues_df = pd.DataFrame(issues, columns=["file", "row", "sku_full", "base_sku", "reason"])
    return out_bytes, issues_df


# =========================
# UI
# =========================
st.set_page_config(page_title="Web App Update Harga", layout="wide")
st.title("Web App Update Harga")

col1, col2, col3 = st.columns(3)

with col1:
    mass_uploads = st.file_uploader(
        "Upload Mass Update (bisa banyak)",
        type=["xlsx"],
        accept_multiple_files=True,
    )

with col2:
    pl_upload = st.file_uploader(
        "Upload Pricelist",
        type=["xlsx", "XLSX"],
        accept_multiple_files=False,
    )

with col3:
    addon_upload = st.file_uploader(
        "Upload Addon Mapping (optional)",
        type=["xlsx", "XLSX"],
        accept_multiple_files=False,
    )

diskon_rp = st.number_input(
    "Diskon (Rp) - mengurangi harga final",
    min_value=0,
    value=0,
    step=1000,
)

run = st.button("Proses")

if run:
    if not mass_uploads:
        st.error("Mass Update wajib diupload.")
        st.stop()
    if not pl_upload:
        st.error("Pricelist wajib diupload.")
        st.stop()

    try:
        mass_files = [(f.name, f.read()) for f in mass_uploads]
        pl_bytes = pl_upload.read()
        addon_bytes = addon_upload.read() if addon_upload else None

        out_bytes, issues_df = process_shopee_files(
            mass_files=mass_files,
            pl_bytes=pl_bytes,
            addon_bytes=addon_bytes,
            diskon_rp=int(diskon_rp),
        )

        st.success("Selesai proses.")

        st.download_button(
            "Download hasil (XLSX)",
            data=out_bytes,
            file_name="hasil_update_shopee.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if issues_df is not None and len(issues_df) > 0:
            st.subheader("Issues Report (tampilan saja)")
            st.dataframe(issues_df, use_container_width=True)

    except Exception as e:
        st.error(str(e))
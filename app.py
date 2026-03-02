import re
import copy
from io import BytesIO
from typing import Dict, Tuple, List, Optional, Set

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


# ============================================================
# Helpers
# ============================================================
AREA_CODE_RE = re.compile(r"^\d+[A-Z]$")  # 0A, 12A, 3B, 19S, etc.


def _norm_str(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


def _norm_key(x) -> str:
    return _norm_str(x).strip().upper()


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


def to_int_or_none(v) -> Optional[int]:
    """Parse stok cell: accept int/float/'1.234'/'1,234' etc."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return int(v)
    if isinstance(v, float):
        if v != v:  # NaN
            return None
        return int(round(v))
    s = _norm_str(v)
    if not s:
        return None
    digits = re.findall(r"\d+", s)
    if not digits:
        return None
    return int("".join(digits))


def find_row_contains(ws: Worksheet, needle: str, scan_rows: int = 200) -> Optional[int]:
    needle_u = needle.upper()
    for r in range(1, min(ws.max_row, scan_rows) + 1):
        for c in range(1, ws.max_column + 1):
            v = _norm_str(ws.cell(r, c).value).upper()
            if v and (needle_u == v or needle_u in v):
                return r
    return None


def map_headers(ws: Worksheet, header_row: int) -> Dict[str, int]:
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        key = _norm_key(v)
        if key:
            headers[key] = c
    return headers


def get_first_sheet(wb: openpyxl.Workbook) -> Worksheet:
    return wb[wb.sheetnames[0]]


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


def get_merged_value(ws: Worksheet, row: int, col: int):
    """Return value of cell; if empty but inside merged range, return top-left value of that merge."""
    cell = ws.cell(row, col)
    if cell.value not in (None, ""):
        return cell.value

    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            return ws.cell(mr.min_row, mr.min_col).value
    return None


def sheet_range_between(sheetnames: List[str], start: str, end: str) -> List[str]:
    """Return inclusive range of sheet names between start and end in workbook order."""
    up = [s.upper() for s in sheetnames]
    if start.upper() not in up or end.upper() not in up:
        raise ValueError(f"Sheet range tidak valid. Pastikan ada '{start}' dan '{end}'.")
    i0 = up.index(start.upper())
    i1 = up.index(end.upper())
    if i0 > i1:
        i0, i1 = i1, i0
    return sheetnames[i0 : i1 + 1]


# ============================================================
# Pricelist processing
# ============================================================
def delete_coming_block_in_laptop(ws: Worksheet):
    """Delete rows from 'COMING' to 'END COMING' inclusive (first occurrence)."""
    r_start = find_row_contains(ws, "COMING", scan_rows=600)
    r_end = find_row_contains(ws, "END COMING", scan_rows=1200)
    if r_start and r_end and r_end >= r_start:
        ws.delete_rows(r_start, r_end - r_start + 1)


def find_sku_header_row(ws: Worksheet) -> int:
    """
    Find a row that contains SKU column name. We'll accept 'SKU NO' or 'KODEBARANG'/'KODE BARANG'.
    """
    candidates = ["SKU NO", "KODEBARANG", "KODE BARANG"]
    for needle in candidates:
        r = find_row_contains(ws, needle, scan_rows=150)
        if r:
            return r
    return 1


def detect_area_and_toko_rows(ws: Worksheet, base_header_row: int) -> Tuple[int, int]:
    """
    Detect which row is AREA and which is TOKO.
    We'll test candidate row pairs and pick the row with more AREA_CODE hits as AREA row.
    """
    pairs = []
    if ws.max_row >= 4:
        pairs.append((3, 4))  # common in your file
    r1 = base_header_row + 1
    r2 = base_header_row + 2
    if ws.max_row >= r2:
        pairs.append((r1, r2))

    def score_area_row(r: int) -> int:
        cnt = 0
        for c in range(1, ws.max_column + 1):
            v = _norm_str(get_merged_value(ws, r, c)).upper()
            if v and AREA_CODE_RE.match(v):
                cnt += 1
        return cnt

    best = None
    best_score = -1
    for a, b in pairs:
        sa = score_area_row(a)
        sb = score_area_row(b)
        score = max(sa, sb)
        if score > best_score:
            best_score = score
            if sa >= sb:
                best = (a, b)  # a is AREA, b is TOKO
            else:
                best = (b, a)  # b is AREA, a is TOKO

    if not best:
        # fallback
        return 4, 3
    return best  # (area_row, toko_row)


def build_stock_lookup_from_sheet(
    ws: Worksheet,
    sheet_name: str,
    keep_area_suffix_for_toko: Set[str],
) -> Tuple[Dict[str, Dict], Set[str], Set[str]]:
    """
    Return:
      sku_map[simple_sku] = {
         "TOT": int|None,
         "by_toko_area": {(TOKO, AREA): int},
      }
    Also return: all_areas, all_toko
    """
    header_row = find_sku_header_row(ws)
    hdr = map_headers(ws, header_row)

    sku_col = hdr.get("SKU NO") or hdr.get("KODEBARANG") or hdr.get("KODE BARANG")
    if not sku_col:
        sku_col = 1

    # Find TOT column: prefer header_row; fallback scan top 12 rows
    tot_col = None
    for c in range(1, ws.max_column + 1):
        v = _norm_key(ws.cell(header_row, c).value)
        if v == "TOT":
            tot_col = c
            break
    if not tot_col:
        for r in range(1, min(12, ws.max_row) + 1):
            for c in range(1, ws.max_column + 1):
                v = _norm_key(ws.cell(r, c).value)
                if v == "TOT":
                    tot_col = c
                    header_row = r
                    break
            if tot_col:
                break
    if not tot_col:
        raise ValueError(f"[{sheet_name}] Kolom 'TOT' tidak ketemu.")

    area_row, toko_row = detect_area_and_toko_rows(ws, header_row)

    all_areas: Set[str] = set()
    all_toko: Set[str] = set()
    sku_map: Dict[str, Dict] = {}

    def norm_area(area_raw: str, toko: str) -> str:
        a = _norm_str(area_raw).upper()
        t = _norm_str(toko).upper()
        if not a:
            return ""
        # default: drop trailing 'A' except toko tertentu (misal JKT)
        if a.endswith("A") and t not in keep_area_suffix_for_toko:
            return a[:-1]  # 2A -> 2
        return a

    for r in range(header_row + 1, ws.max_row + 1):
        sku_raw = ws.cell(r, sku_col).value
        sku = _norm_str(sku_raw)
        if not sku:
            continue

        sku_key = _norm_key(sku)
        if sku_key in ("TOTAL", "SKU NO"):
            continue

        tot_val = to_int_or_none(ws.cell(r, tot_col).value)
        by_toko_area = {}

        for c in range(tot_col + 1, ws.max_column + 1):
            area_raw = get_merged_value(ws, area_row, c)
            toko_raw = get_merged_value(ws, toko_row, c)

            area_s = _norm_str(area_raw).upper()
            toko_s = _norm_str(toko_raw).upper()

            if not area_s and not toko_s:
                continue

            # Safety swap if a column looks inverted
            if AREA_CODE_RE.match(toko_s) and not AREA_CODE_RE.match(area_s):
                area_s, toko_s = toko_s, area_s

            if not area_s or not toko_s:
                continue
            if not AREA_CODE_RE.match(area_s):
                # If area cell isn't a code, skip this column
                continue

            area_n = norm_area(area_s, toko_s)
            if not area_n:
                continue

            v = to_int_or_none(ws.cell(r, c).value)
            if v is None:
                continue

            all_areas.add(area_n)
            all_toko.add(toko_s)
            by_toko_area[(toko_s, area_n)] = v

        sku_map[sku_key] = {"TOT": tot_val, "by_toko_area": by_toko_area}

    return sku_map, all_areas, all_toko


def build_stock_lookup_from_pricelist(
    pl_bytes: bytes,
    keep_area_suffix_for_toko: Set[str],
) -> Tuple[Dict[str, Dict], Set[str], Set[str]]:
    """
    Rules:
      - delete COMING..END COMING in LAPTOP
      - only read sheets LAPTOP..SER OTH CON
    """
    wb = openpyxl.load_workbook(BytesIO(pl_bytes), data_only=True)

    # Clean COMING block in LAPTOP
    for s in wb.sheetnames:
        if s.upper() == "LAPTOP":
            delete_coming_block_in_laptop(wb[s])
            break

    target_sheets = sheet_range_between(wb.sheetnames, "LAPTOP", "SER OTH CON")

    merged_map: Dict[str, Dict] = {}
    all_areas: Set[str] = set()
    all_toko: Set[str] = set()

    for s in target_sheets:
        ws = wb[s]
        sku_map, areas, toko = build_stock_lookup_from_sheet(ws, s, keep_area_suffix_for_toko)
        # If duplicates, later sheet overwrites (workbook order)
        merged_map.update(sku_map)
        all_areas |= areas
        all_toko |= toko

    if not merged_map:
        raise ValueError("Pricelist terbaca, tapi lookup stok kosong.")
    return merged_map, all_areas, all_toko


# ============================================================
# Mass update processing (stok)
# ============================================================
def detect_mass_columns(ws: Worksheet) -> Tuple[int, int, int]:
    """
    Detect:
      data_start_row: default 7
      sku_col: header contains 'SKU'
      stock_col: PRIORITAS 'KUANTITAS', fallback STOK/STOCK/QTY
    """
    data_start = 7
    header_scan_rows = [1, 2, 3, 4, 5, 6]
    best_sku_col = None
    best_stock_col = None

    for r in header_scan_rows:
        for c in range(1, ws.max_column + 1):
            v = _norm_key(ws.cell(r, c).value)
            if not v:
                continue

            if best_sku_col is None and "SKU" in v:
                best_sku_col = c

            if best_stock_col is None and "KUANTITAS" in v:
                best_stock_col = c

            if best_stock_col is None and (("STOK" in v) or ("STOCK" in v) or (v == "QTY") or ("QTY" in v)):
                best_stock_col = c

    if best_sku_col is None:
        best_sku_col = 6  # fallback
    if best_stock_col is None:
        best_stock_col = 8  # fallback, can override
    return data_start, best_sku_col, best_stock_col


def pick_stock_value(
    sku_full: str,
    stock_lookup: Dict[str, Dict],
    mode: str,
    chosen_areas: Set[str],
    chosen_toko: Set[str],
    chosen_pairs: Set[Tuple[str, str]],
) -> Optional[int]:
    """
    SKU with +ADDON => use BASE SKU stock (before '+')
    """
    base, _addons = split_sku_addons(sku_full)
    base_key = _norm_key(base)
    if not base_key or base_key not in stock_lookup:
        return None

    rec = stock_lookup[base_key]
    tot = rec.get("TOT")
    by_toko_area: Dict[Tuple[str, str], int] = rec.get("by_toko_area", {}) or {}

    if mode == "National (TOT)":
        return tot if tot is not None else None

    if mode == "Per AREA (sum)":
        if not chosen_areas:
            return None
        s = 0
        hit = False
        for (t, a), v in by_toko_area.items():
            if a in chosen_areas:
                s += int(v)
                hit = True
        return s if hit else None

    if mode == "Per TOKO (sum)":
        if not chosen_toko:
            return None
        s = 0
        hit = False
        for (t, a), v in by_toko_area.items():
            if t in chosen_toko:
                s += int(v)
                hit = True
        return s if hit else None

    if mode == "Per TOKO+AREA (sum)":
        if not chosen_pairs:
            return None
        s = 0
        hit = False
        for key, v in by_toko_area.items():
            if key in chosen_pairs:
                s += int(v)
                hit = True
        return s if hit else None

    return None


def process_mass_update_stock(
    mass_files: List[Tuple[str, bytes]],
    pl_bytes: bytes,
    keep_area_suffix_for_toko: Set[str],
    mode: str,
    chosen_areas: Set[str],
    chosen_toko: Set[str],
    chosen_pairs: Set[Tuple[str, str]],
    override_sku_col: Optional[int] = None,
    override_stock_col: Optional[int] = None,
    data_start_row_override: Optional[int] = None,
) -> Tuple[bytes, pd.DataFrame]:
    """
    Output one xlsx bytes:
      - template from first mass file
      - rows from data_start only include updated rows
      - only stock column overwritten
    """
    stock_lookup, _, _ = build_stock_lookup_from_pricelist(pl_bytes, keep_area_suffix_for_toko)

    issues = []

    first_name, first_bytes = mass_files[0]
    out_wb = openpyxl.load_workbook(BytesIO(first_bytes))
    out_ws = get_first_sheet(out_wb)

    data_start, sku_col, stock_col = detect_mass_columns(out_ws)
    if data_start_row_override:
        data_start = int(data_start_row_override)
    if override_sku_col:
        sku_col = int(override_sku_col)
    if override_stock_col:
        stock_col = int(override_stock_col)

    # Clear existing data rows
    if out_ws.max_row >= data_start:
        out_ws.delete_rows(data_start, out_ws.max_row - data_start + 1)

    # Style template row
    tmp_wb = openpyxl.load_workbook(BytesIO(first_bytes))
    tmp_ws = get_first_sheet(tmp_wb)
    template_style_row = data_start
    max_col = tmp_ws.max_column

    write_row = data_start
    total_updated = 0

    for fname, fbytes in mass_files:
        try:
            wb = openpyxl.load_workbook(BytesIO(fbytes))
            ws = get_first_sheet(wb)

            for r in range(data_start, ws.max_row + 1):
                sku_full = _norm_str(ws.cell(r, sku_col).value)
                if not sku_full:
                    continue

                old_stock = to_int_or_none(ws.cell(r, stock_col).value)

                new_stock = pick_stock_value(
                    sku_full=sku_full,
                    stock_lookup=stock_lookup,
                    mode=mode,
                    chosen_areas=chosen_areas,
                    chosen_toko=chosen_toko,
                    chosen_pairs=chosen_pairs,
                )
                if new_stock is None:
                    continue

                if old_stock is not None and int(old_stock) == int(new_stock):
                    continue

                out_ws.insert_rows(write_row, 1)
                copy_row_style(tmp_ws, template_style_row, out_ws, write_row, max_col)

                # Copy all values from source row
                for c in range(1, max_col + 1):
                    out_ws.cell(write_row, c).value = ws.cell(r, c).value

                # Overwrite ONLY stok cell
                out_ws.cell(write_row, stock_col).value = int(new_stock)

                total_updated += 1
                write_row += 1

        except Exception as e:
            issues.append({"file": fname, "row": "", "sku_full": "", "reason": f"Gagal proses file: {e}"})

    if total_updated == 0:
        issues.append(
            {"file": "", "row": "", "sku_full": "", "reason": "Tidak ada baris yang berubah / tidak ada SKU yang match."}
        )

    out_buf = BytesIO()
    out_wb.save(out_buf)
    out_bytes = out_buf.getvalue()

    issues_df = pd.DataFrame(issues, columns=["file", "row", "sku_full", "reason"])
    return out_bytes, issues_df


# ============================================================
# UI (Streamlit)
# ============================================================
st.set_page_config(page_title="Update Stok Mass", layout="wide")
st.title("Update Stok (Mass Update)")

col1, col2, col3 = st.columns(3)

with col1:
    mass_uploads = st.file_uploader(
        "Upload Mass Update (bisa banyak)",
        type=["xlsx"],
        accept_multiple_files=True,
    )

with col2:
    pl_upload = st.file_uploader(
        "Upload Pricelist (multi-sheet)",
        type=["xlsx", "XLSX"],
        accept_multiple_files=False,
    )

with col3:
    _addon_upload = st.file_uploader(
        "Upload Addon Mapping (optional) - TIDAK dipakai untuk stok",
        type=["xlsx", "XLSX"],
        accept_multiple_files=False,
    )

st.caption("Catatan: SKU yang mengandung '+ADDON' akan pakai stok BASE SKU (sebelum '+').")

with st.expander("Rule khusus area suffix (default: JKT mempertahankan 3B/3C/..)", expanded=True):
    keep_toko_text = st.text_input(
        "TOKO yang mempertahankan suffix area (misal JKT)",
        value="JKT",
        help="Pisahkan dengan koma. Contoh: JKT,SUB",
    )
    keep_area_suffix_for_toko = {x.strip().upper() for x in keep_toko_text.split(",") if x.strip()}

mode = st.radio(
    "Pilih sumber stok untuk update",
    options=["National (TOT)", "Per AREA (sum)", "Per TOKO (sum)", "Per TOKO+AREA (sum)"],
    horizontal=True,
)

areas_selected: Set[str] = set()
toko_selected: Set[str] = set()
pairs_selected: Set[Tuple[str, str]] = set()

all_areas = []
all_toko = []
if pl_upload is not None:
    try:
        _, areas_preview, toko_preview = build_stock_lookup_from_pricelist(
            pl_upload.getvalue(),
            keep_area_suffix_for_toko,
        )
        all_areas = sorted(list(areas_preview))
        all_toko = sorted(list(toko_preview))
    except Exception as e:
        st.error(f"Gagal baca pricelist untuk daftar AREA/TOKO: {e}")

if mode == "Per AREA (sum)":
    areas_selected = set(st.multiselect("Pilih AREA (boleh banyak, akan dijumlah)", options=all_areas))
elif mode == "Per TOKO (sum)":
    toko_selected = set(st.multiselect("Pilih TOKO (boleh banyak, akan dijumlah)", options=all_toko))
elif mode == "Per TOKO+AREA (sum)":
    left, right = st.columns(2)
    with left:
        toko_selected = set(st.multiselect("Pilih TOKO", options=all_toko))
    with right:
        areas_selected = set(st.multiselect("Pilih AREA", options=all_areas))

    pair_labels = []
    pair_map = {}
    for t in sorted(toko_selected):
        for a in sorted(areas_selected):
            label = f"{t} | {a}"
            pair_labels.append(label)
            pair_map[label] = (t, a)

    chosen_labels = st.multiselect(
        "Pilih kombinasi TOKO|AREA (boleh banyak, akan dijumlah)",
        options=pair_labels,
    )
    pairs_selected = {pair_map[x] for x in chosen_labels}

with st.expander("Advanced (kalau kolom SKU / KUANTITAS di Mass Update beda)", expanded=False):
    sku_col_override = st.number_input("SKU column index (1=A)", min_value=0, value=0, step=1)
    stock_col_override = st.number_input("KUANTITAS column index (1=A)", min_value=0, value=0, step=1)
    data_start_override = st.number_input("Data mulai row", min_value=0, value=0, step=1)
    st.caption("Kalau 0 berarti auto-detect/default. Umumnya data mulai row 7.")

run = st.button("Proses Update Stok")

if run:
    if not mass_uploads:
        st.error("Mass Update wajib diupload.")
        st.stop()
    if not pl_upload:
        st.error("Pricelist wajib diupload.")
        st.stop()

    if mode == "Per AREA (sum)" and not areas_selected:
        st.error("Mode Per AREA: pilih minimal 1 AREA.")
        st.stop()
    if mode == "Per TOKO (sum)" and not toko_selected:
        st.error("Mode Per TOKO: pilih minimal 1 TOKO.")
        st.stop()
    if mode == "Per TOKO+AREA (sum)" and not pairs_selected:
        st.error("Mode Per TOKO+AREA: pilih minimal 1 kombinasi TOKO|AREA.")
        st.stop()

    try:
        mass_files = [(f.name, f.read()) for f in mass_uploads]
        pl_bytes = pl_upload.getvalue()

        out_bytes, issues_df = process_mass_update_stock(
            mass_files=mass_files,
            pl_bytes=pl_bytes,
            keep_area_suffix_for_toko=keep_area_suffix_for_toko,
            mode=mode,
            chosen_areas=areas_selected,
            chosen_toko=toko_selected,
            chosen_pairs=pairs_selected,
            override_sku_col=int(sku_col_override) if sku_col_override else None,
            override_stock_col=int(stock_col_override) if stock_col_override else None,
            data_start_row_override=int(data_start_override) if data_start_override else None,
        )

        st.success("Selesai proses update stok.")
        st.download_button(
            "Download hasil (XLSX)",
            data=out_bytes,
            file_name="hasil_update_stok.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if issues_df is not None and len(issues_df) > 0:
            st.subheader("Issues Report (tampilan saja)")
            st.dataframe(issues_df, use_container_width=True)

    except Exception as e:
        st.error(str(e))

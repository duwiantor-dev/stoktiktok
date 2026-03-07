import re
from io import BytesIO
from typing import Dict, Tuple, List, Optional, Set

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


# ============================================================
# CONFIG
# ============================================================
MAX_MASS_FILES = 50
MAX_TOTAL_UPLOAD_MB = 200


# ============================================================
# HELPERS
# ============================================================
def _norm_str(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


def norm_sku(v) -> str:
    s = _norm_str(v).upper()
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"\s+", "", s)
    return s


def split_sku_addons(sku_full: str) -> Tuple[str, List[str]]:
    parts = [p for p in _norm_str(sku_full).split("+") if p.strip()]
    if not parts:
        return "", []
    return parts[0].strip(), [p.strip() for p in parts[1:]]


def to_int_or_none(v) -> Optional[int]:
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return int(v)
    if isinstance(v, float):
        if v != v:
            return None
        return int(round(v))
    s = _norm_str(v)
    if not s:
        return None
    digits = re.findall(r"\d+", s)
    if not digits:
        return None
    return int("".join(digits))


def get_first_sheet(wb: openpyxl.Workbook) -> Worksheet:
    return wb[wb.sheetnames[0]]


def find_row_contains(ws: Worksheet, needle: str, scan_rows: int = 300) -> Optional[int]:
    needle_u = needle.upper()
    for r in range(1, min(ws.max_row, scan_rows) + 1):
        for c in range(1, ws.max_column + 1):
            v = _norm_str(ws.cell(r, c).value).upper()
            if v and (needle_u == v or needle_u in v):
                return r
    return None


def sheet_range_between(sheetnames: List[str], start: str, end: str) -> List[str]:
    up = [s.upper() for s in sheetnames]
    if start.upper() not in up or end.upper() not in up:
        raise ValueError(f"Sheet range tidak valid. Pastikan ada '{start}' dan '{end}'.")
    i0 = up.index(start.upper())
    i1 = up.index(end.upper())
    if i0 > i1:
        i0, i1 = i1, i0
    return sheetnames[i0:i1 + 1]


def norm_area_name(area_raw) -> str:
    return _norm_str(area_raw).upper()


def total_upload_size_mb(files: List) -> float:
    total = 0
    for f in files:
        try:
            total += len(f.getvalue())
        except Exception:
            pass
    return total / (1024 * 1024)


# ============================================================
# PRICE LIST PARSING (TOT + AREA ONLY)
# ============================================================
def delete_coming_block_in_laptop(ws: Worksheet):
    r_start = find_row_contains(ws, "COMING", scan_rows=600)
    r_end = find_row_contains(ws, "END COMING", scan_rows=1200)
    if r_start and r_end and r_end >= r_start:
        ws.delete_rows(r_start, r_end - r_start + 1)


def find_header_row_by_exact(ws: Worksheet, header_text: str, scan_rows: int = 150) -> Optional[int]:
    target = header_text.strip().upper()
    for r in range(1, min(ws.max_row, scan_rows) + 1):
        for c in range(1, ws.max_column + 1):
            v = _norm_str(ws.cell(r, c).value).strip().upper()
            if v == target:
                return r
    return None


def find_tot_col(ws: Worksheet, header_row_hint: int) -> Tuple[int, int]:
    for c in range(1, ws.max_column + 1):
        if _norm_str(ws.cell(header_row_hint, c).value).strip().upper() == "TOT":
            return header_row_hint, c

    for r in range(1, min(12, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            if _norm_str(ws.cell(r, c).value).strip().upper() == "TOT":
                return r, c

    raise ValueError("Kolom 'TOT' tidak ketemu.")


def build_merged_lookup_map(ws: Worksheet) -> Dict[Tuple[int, int], object]:
    merged_map: Dict[Tuple[int, int], object] = {}
    for mr in ws.merged_cells.ranges:
        top_left_val = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merged_map[(r, c)] = top_left_val
    return merged_map


def get_cell_or_merged_value(ws: Worksheet, merged_map: Dict[Tuple[int, int], object], row: int, col: int):
    v = ws.cell(row, col).value
    if v not in (None, ""):
        return v
    return merged_map.get((row, col))


def build_area_meta(
    ws: Worksheet,
    merged_map: Dict[Tuple[int, int], object],
    area_row: int,
    start_col: int,
) -> Dict[int, str]:
    col_area: Dict[int, str] = {}

    for c in range(start_col, ws.max_column + 1):
        area_raw = get_cell_or_merged_value(ws, merged_map, area_row, c)
        area_name = norm_area_name(area_raw)
        if not area_name:
            continue
        col_area[c] = area_name

    return col_area


def build_stock_lookup_from_sheet_fast(ws: Worksheet, sheet_name: str):
    AREA_ROW = 3

    header_row = find_header_row_by_exact(ws, "KODEBARANG", scan_rows=150)
    if header_row is None:
        header_row = find_header_row_by_exact(ws, "KODE BARANG", scan_rows=150)
    if header_row is None:
        raise ValueError(f"[{sheet_name}] Header 'KODEBARANG' tidak ketemu.")

    sku_col = None
    for c in range(1, ws.max_column + 1):
        v = _norm_str(ws.cell(header_row, c).value).strip().upper()
        if v in ("KODEBARANG", "KODE BARANG"):
            sku_col = c
            break

    if sku_col is None:
        raise ValueError(f"[{sheet_name}] Kolom 'KODEBARANG' / 'KODE BARANG' tidak ditemukan.")

    header_row_used, tot_col = find_tot_col(ws, header_row)

    merged_map = build_merged_lookup_map(ws)
    col_area = build_area_meta(
        ws=ws,
        merged_map=merged_map,
        area_row=AREA_ROW,
        start_col=tot_col + 1,
    )

    sku_map: Dict[str, Dict] = {}
    areas: Set[str] = set(col_area.values())

    for r in range(header_row_used + 1, ws.max_row + 1):
        sku_raw = ws.cell(r, sku_col).value
        sku = _norm_str(sku_raw)
        if not sku:
            continue

        sku_key = norm_sku(sku)
        if sku_key in ("TOTAL", "KODEBARANG", "KODE BARANG", "KODEBARANG."):
            continue

        tot_val = to_int_or_none(ws.cell(r, tot_col).value)
        by_area: Dict[str, int] = {}

        for c, area_name in col_area.items():
            v = to_int_or_none(ws.cell(r, c).value)
            if v is None:
                continue
            by_area[area_name] = by_area.get(area_name, 0) + int(v)

        sku_map[sku_key] = {
            "TOT": tot_val,
            "by_area": by_area,
        }

    return sku_map, sorted(areas)


@st.cache_data(show_spinner=False)
def build_stock_lookup_from_pricelist_cached(pl_bytes: bytes):
    wb = openpyxl.load_workbook(BytesIO(pl_bytes), data_only=True, read_only=False)

    for s in wb.sheetnames:
        if s.upper() == "LAPTOP":
            delete_coming_block_in_laptop(wb[s])
            break

    target_sheets = sheet_range_between(wb.sheetnames, "LAPTOP", "SER OTH CON")

    merged_lookup: Dict[str, Dict] = {}
    areas_all: Set[str] = set()

    for s in target_sheets:
        ws = wb[s]
        sku_map, areas = build_stock_lookup_from_sheet_fast(ws, s)
        merged_lookup.update(sku_map)
        areas_all |= set(areas)

    if not merged_lookup:
        raise ValueError("Pricelist terbaca, tapi lookup stok kosong.")

    return merged_lookup, sorted(areas_all)


# ============================================================
# TIKTOK MASS UPDATE LAYOUT
# ============================================================
def find_tiktok_columns_normal(ws: Worksheet) -> Tuple[int, int, int]:
    HEADER_ROW = 3
    DATA_START_ROW = 6

    sku_col = None
    qty_col = None

    for c in range(1, ws.max_column + 1):
        v = _norm_str(ws.cell(HEADER_ROW, c).value).strip().upper()
        if v in ("SKU PENJUAL", "SELLER SKU"):
            sku_col = c
        if v == "KUANTITAS":
            qty_col = c

    if not sku_col:
        raise ValueError("Kolom SKU tidak ketemu. Pastikan header 'SKU Penjual' / 'Seller SKU' ada di row 3.")
    if not qty_col:
        raise ValueError("Kolom Kuantitas tidak ketemu. Pastikan header 'Kuantitas' ada di row 3.")

    return DATA_START_ROW, sku_col, qty_col


def find_tiktok_columns_readonly(ws) -> Tuple[int, int, int]:
    HEADER_ROW = 3
    DATA_START_ROW = 6

    sku_col = None
    qty_col = None

    row_vals = list(ws.iter_rows(min_row=HEADER_ROW, max_row=HEADER_ROW, values_only=True))[0]
    for idx, val in enumerate(row_vals, start=1):
        v = _norm_str(val).strip().upper()
        if v in ("SKU PENJUAL", "SELLER SKU"):
            sku_col = idx
        if v == "KUANTITAS":
            qty_col = idx

    if not sku_col:
        raise ValueError("Kolom SKU tidak ketemu. Pastikan header 'SKU Penjual' / 'Seller SKU' ada di row 3.")
    if not qty_col:
        raise ValueError("Kolom Kuantitas tidak ketemu. Pastikan header 'Kuantitas' ada di row 3.")

    return DATA_START_ROW, sku_col, qty_col


def pick_stock_value(
    sku_full: str,
    stock_lookup: Dict[str, Dict],
    mode: str,
    chosen_areas: Set[str],
) -> Optional[int]:
    base, _ = split_sku_addons(sku_full)
    base_key = norm_sku(base)

    if not base_key or base_key not in stock_lookup:
        return None

    rec = stock_lookup[base_key]
    tot = rec.get("TOT")
    by_area: Dict[str, int] = rec.get("by_area", {}) or {}

    if mode == "Stok Nasional (TOT)":
        return tot if tot is not None else None

    if mode == "Stok Area":
        if not chosen_areas:
            return None

        s = 0
        hit = False
        for area_name, v in by_area.items():
            if area_name in chosen_areas:
                s += int(v)
                hit = True

        return s if hit else None

    return None


# ============================================================
# PROCESSOR
# ============================================================
def collect_changed_rows_from_mass_file(
    file_name: str,
    file_bytes: bytes,
    stock_lookup: Dict[str, Dict],
    mode: str,
    chosen_areas: Set[str],
):
    stats = {
        "rows_scanned": 0,
        "rows_written": 0,
        "rows_unchanged": 0,
        "rows_unmatched": 0,
    }

    changed_rows = []

    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=False)
    ws = wb[wb.sheetnames[0]]

    data_start, sku_col, qty_col = find_tiktok_columns_readonly(ws)

    for row in ws.iter_rows(min_row=data_start, values_only=True):
        row_list = list(row)
        if not row_list:
            continue

        sku_full = _norm_str(row_list[sku_col - 1] if len(row_list) >= sku_col else None)
        if not sku_full:
            continue

        stats["rows_scanned"] += 1

        old_qty = to_int_or_none(row_list[qty_col - 1] if len(row_list) >= qty_col else None)
        new_qty = pick_stock_value(
            sku_full=sku_full,
            stock_lookup=stock_lookup,
            mode=mode,
            chosen_areas=chosen_areas,
        )

        if new_qty is None:
            stats["rows_unmatched"] += 1
            continue

        if old_qty is not None and int(old_qty) == int(new_qty):
            stats["rows_unchanged"] += 1
            continue

        if len(row_list) < qty_col:
            row_list.extend([None] * (qty_col - len(row_list)))

        row_list[qty_col - 1] = int(new_qty)
        changed_rows.append(row_list)
        stats["rows_written"] += 1

    wb.close()
    return changed_rows, stats


def write_output_from_template(template_bytes: bytes, changed_rows_all: List[List[object]]) -> bytes:
    out_wb = openpyxl.load_workbook(BytesIO(template_bytes))
    out_ws = get_first_sheet(out_wb)

    data_start, _, _ = find_tiktok_columns_normal(out_ws)

    if out_ws.max_row >= data_start:
        out_ws.delete_rows(data_start, out_ws.max_row - data_start + 1)

    for idx, row_vals in enumerate(changed_rows_all, start=data_start):
        for c, val in enumerate(row_vals, start=1):
            out_ws.cell(idx, c).value = val

    buf = BytesIO()
    out_wb.save(buf)
    return buf.getvalue()


def process_mass_update_stock_tiktok_fast(
    mass_files: List[Tuple[str, bytes]],
    stock_lookup: Dict[str, Dict],
    mode: str,
    chosen_areas: Set[str],
    progress_slot=None,
) -> Tuple[bytes, pd.DataFrame, Dict[str, int]]:
    issues = []
    changed_rows_all: List[List[object]] = []

    stats_total = {
        "files_total": len(mass_files),
        "rows_scanned": 0,
        "rows_written": 0,
        "rows_unchanged": 0,
        "rows_unmatched": 0,
    }

    for idx, (fname, fbytes) in enumerate(mass_files, start=1):
        try:
            changed_rows, stats = collect_changed_rows_from_mass_file(
                file_name=fname,
                file_bytes=fbytes,
                stock_lookup=stock_lookup,
                mode=mode,
                chosen_areas=chosen_areas,
            )
            changed_rows_all.extend(changed_rows)

            for k in ("rows_scanned", "rows_written", "rows_unchanged", "rows_unmatched"):
                stats_total[k] += stats[k]

        except Exception as e:
            issues.append({"file": fname, "reason": f"Gagal proses file: {e}"})

        if progress_slot is not None:
            progress_slot.progress(idx / len(mass_files), text=f"Memproses file {idx}/{len(mass_files)}")

    if stats_total["rows_written"] == 0:
        issues.append({"file": "", "reason": "Tidak ada baris berubah / tidak ada SKU yang match."})

    out_bytes = write_output_from_template(mass_files[0][1], changed_rows_all)
    issues_df = pd.DataFrame(issues, columns=["file", "reason"])

    return out_bytes, issues_df, stats_total


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="Update Stok TikTok (Mass Update)", layout="wide")
st.title("Update Stok TikTok (Mass Update)")

if "stock_lookup" not in st.session_state:
    st.session_state.stock_lookup = None
if "areas" not in st.session_state:
    st.session_state.areas = []
if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None
if "issues_df" not in st.session_state:
    st.session_state.issues_df = None
if "stats" not in st.session_state:
    st.session_state.stats = None
if "last_result_name" not in st.session_state:
    st.session_state.last_result_name = "hasil_update_stok_tiktok.xlsx"

col1, col2 = st.columns(2)

with col1:
    mass_uploads = st.file_uploader(
        "Upload Mass Update TikTok (bisa banyak)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="mass_uploads",
    )

with col2:
    pl_upload = st.file_uploader(
        "Upload Pricelist (multi-sheet)",
        type=["xlsx", "XLSX"],
        accept_multiple_files=False,
        key="pl_upload",
    )

st.caption("Catatan: SKU yang mengandung '+ADDON' akan pakai stok BASE SKU (sebelum '+').")

if mass_uploads:
    if len(mass_uploads) > MAX_MASS_FILES:
        st.error(f"Maksimal {MAX_MASS_FILES} file mass update per proses.")
        st.stop()

    total_mb = total_upload_size_mb(mass_uploads)
    if total_mb > MAX_TOTAL_UPLOAD_MB:
        st.error(f"Total ukuran file terlalu besar: {total_mb:.1f} MB. Maksimal {MAX_TOTAL_UPLOAD_MB} MB.")
        st.stop()

load_btn = st.button("Load Data Pricelist", type="secondary", key="btn_load_data")

if load_btn:
    if not pl_upload:
        st.error("Upload Pricelist dulu.")
        st.stop()

    try:
        with st.spinner("Membaca pricelist..."):
            pl_bytes = pl_upload.getvalue()
            lookup, areas = build_stock_lookup_from_pricelist_cached(pl_bytes)
            st.session_state.stock_lookup = lookup
            st.session_state.areas = areas

        st.success(f"OK. Ditemukan {len(areas)} AREA.")
    except Exception as e:
        st.error(f"Pricelist tidak valid: {e}")
        st.stop()

mode = None
chosen_areas: Set[str] = set()
can_show_run_button = False

if st.session_state.stock_lookup is None:
    st.info("Klik 'Load Data Pricelist' dulu supaya daftar AREA muncul.")
else:
    mode = st.radio(
        "Pilih sumber stok untuk update",
        options=["Stok Nasional (TOT)", "Stok Area"],
        horizontal=True,
        key="mode_stock_source",
    )

    if mode == "Stok Area":
        chosen_areas = set(
            st.multiselect(
                "Pilih AREA (boleh banyak)",
                options=st.session_state.areas,
                key="ms_areas",
            )
        )
        if chosen_areas:
            can_show_run_button = True
    elif mode == "Stok Nasional (TOT)":
        can_show_run_button = True

run = False
if can_show_run_button:
    run = st.button("Proses Update Stok", type="primary", key="btn_run")

if run:
    if not mass_uploads:
        st.error("Mass Update wajib diupload.")
        st.stop()

    if st.session_state.stock_lookup is None:
        st.error("Klik 'Load Data Pricelist' dulu.")
        st.stop()

    if mode == "Stok Area" and not chosen_areas:
        st.error("Mode Stok Area: pilih minimal 1 AREA.")
        st.stop()

    try:
        progress = st.progress(0, text="Menyiapkan proses...")

        with st.spinner("Sedang proses update stok TikTok..."):
            mass_files = [(f.name, f.getvalue()) for f in mass_uploads]

            out_bytes, issues_df, stats = process_mass_update_stock_tiktok_fast(
                mass_files=mass_files,
                stock_lookup=st.session_state.stock_lookup,
                mode=mode,
                chosen_areas=chosen_areas,
                progress_slot=progress,
            )

            st.session_state.result_bytes = out_bytes
            st.session_state.issues_df = issues_df
            st.session_state.stats = stats
            st.session_state.last_result_name = "hasil_update_stok_tiktok.xlsx"

        progress.progress(1.0, text="Selesai")
        st.success("Selesai proses update stok.")

    except Exception as e:
        st.error(str(e))

if st.session_state.result_bytes is not None:
    st.subheader("Hasil Proses")

    if st.session_state.stats:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("File diproses", st.session_state.stats["files_total"])
        c2.metric("Baris discan", st.session_state.stats["rows_scanned"])
        c3.metric("Baris diupdate", st.session_state.stats["rows_written"])
        c4.metric("SKU tidak match", st.session_state.stats["rows_unmatched"])

    st.download_button(
        "Download hasil (XLSX)",
        data=st.session_state.result_bytes,
        file_name=st.session_state.last_result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_result",
    )

    if st.session_state.issues_df is not None and len(st.session_state.issues_df) > 0:
        st.subheader("Issues Report")
        st.dataframe(st.session_state.issues_df, use_container_width=True)

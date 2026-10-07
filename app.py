import re
import hashlib
from copy import copy
from datetime import date, datetime, timedelta
from io import BytesIO

import streamlit as st
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

st.set_page_config(
    page_title="KD-EXAM",
    page_icon="KD_EXAM_logo.png",
    layout="wide",
)

# KD-EXAM branding: faint pink theme + student logo.
st.markdown(
    """
    <style>
    :root {
        --kd-blue: #1976d2;
    }
    .stApp {
        background: linear-gradient(180deg, #fff8fb 0%, #fffdfd 48%, #fff6fa 100%);
    }
    .kd-brand {
        display: flex;
        align-items: center;
        gap: 18px;
        padding: 14px 18px;
        margin: 0 0 18px 0;
        border: 1px solid #f3c9da;
        border-radius: 22px;
        background: linear-gradient(135deg, rgba(255,255,255,.96), rgba(253,230,240,.92));
        box-shadow: 0 6px 20px rgba(200, 80, 130, .08);
    }
    .kd-brand img {
        width: 82px;
        height: 82px;
        object-fit: cover;
        border-radius: 20px;
        border: 2px solid #f3a9c7;
        background: white;
    }
    .kd-name {
        margin: 0;
        line-height: .95;
        font-size: clamp(34px, 6vw, 58px);
        font-weight: 900;
        letter-spacing: 1px;
        color: #b3135b;
        font-style: italic;
    }
    .kd-name span {
        color: var(--kd-blue);
        font-style: italic;
        font-weight: 900;
        letter-spacing: 2px;
    }
    .kd-tag {
        margin: 8px 0 0;
        color: #7c4660;
        font-size: 15px;
    }
    .kd-version {
        color: #8d6575;
        font-size: 13px;
        margin-top: 3px;
    }
    div.stButton > button, div[data-testid="stDownloadButton"] button {
        border-color: #efb0c9 !important;
        border-radius: 12px !important;
    }
    div.stButton > button:hover, div[data-testid="stDownloadButton"] button:hover {
        border-color: #d94d87 !important;
        color: #b3135b !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

from pathlib import Path
_logo_path = Path(__file__).with_name("KD_EXAM_logo.png")
_logo_b64 = ""
if _logo_path.exists():
    import base64
    _logo_b64 = base64.b64encode(_logo_path.read_bytes()).decode("ascii")

if _logo_b64:
    st.markdown(
        f"""
        <div class="kd-brand">
            <img src="data:image/png;base64,{_logo_b64}" alt="KD-EXAM logo">
            <div>
                <div class="kd-name">KD <span>EXAM</span></div>
                <div class="kd-tag">Study &nbsp;|&nbsp; Practice &nbsp;|&nbsp; Succeed</div>
                <div class="kd-version">Version 5.11 — Date-wise Report, Dashboard/Summary and Excel/PDF Print Report</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.title("KD-EXAM")
    st.caption("Version 5.11 — Date-wise Report, Dashboard/Summary and Excel/PDF Print Report")

TARGET_HEADERS = ["Total", "Block", "Super", "Squad", "Sessions"]


def wb_bytes(wb):
    # Always open generated Excel workbooks at A1 on the main report sheet.
    # This changes only the saved opening view/selection, not report data or calculations.
    try:
        main_ws = get_main_sheet(wb) if 'get_main_sheet' in globals() else wb.active
        main_idx = wb.worksheets.index(main_ws)
        wb.active = main_idx

        for ws in wb.worksheets:
            ws.sheet_view.topLeftCell = "A1"
            if ws.sheet_view.selection:
                sel = ws.sheet_view.selection[0]
                sel.activeCell = "A1"
                sel.sqref = "A1"
            ws.sheet_view.tabSelected = (ws is main_ws)
    except Exception:
        pass

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def norm_header(v):
    return re.sub(r"\s+", " ", str(v or "").strip().lower())


def find_header_row(ws, required, max_rows=20):
    req = {norm_header(x) for x in required}
    for r in range(1, min(ws.max_row, max_rows) + 1):
        vals = {norm_header(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        if req.issubset(vals):
            return r
    return None


def find_col(ws, names, header_row=1):
    wanted = {norm_header(x) for x in names}
    for c in range(1, ws.max_column + 1):
        if norm_header(ws.cell(header_row, c).value) in wanted:
            return c
    return None


def get_main_sheet(wb):
    if " bundle all" in wb.sheetnames:
        return wb[" bundle all"]
    for ws in wb.worksheets:
        h = find_header_row(ws, {"date", "time"})
        if h:
            return ws
    return wb.active


def date_key(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


def display_date(k):
    try:
        return datetime.strptime(k, "%Y-%m-%d").strftime("%d/%m")
    except Exception:
        return str(k)


def norm_time(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%I:%M %p").lstrip("0").lower()
    if hasattr(v, "strftime") and not isinstance(v, str):
        try:
            return v.strftime("%I:%M %p").lstrip("0").lower()
        except Exception:
            pass
    s = str(v).strip().lower().replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+to\s+", "-", s)
    s = s.replace(".30", ":30").replace(".00", ":00")
    s = re.sub(r"\s*([:-])\s*", r"\1", s)
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"(\d)(am|pm)$", r"\1 \2", s)
    return s


def copy_cell_style(src, dst):
    if isinstance(dst, MergedCell) or isinstance(src, MergedCell):
        return
    if src.has_style:
        dst._style = copy(src._style)
    dst.number_format = src.number_format
    dst.alignment = copy(src.alignment)
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.protection = copy(src.protection)


def copy_col_dimension(ws, src_col, dst_col):
    s = ws.column_dimensions[get_column_letter(src_col)]
    d = ws.column_dimensions[get_column_letter(dst_col)]
    d.width = s.width
    d.hidden = s.hidden
    d.bestFit = s.bestFit
    d.outlineLevel = s.outlineLevel
    d.collapsed = s.collapsed


def unmerge_ranges_in_column(ws, col):
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == col and rng.max_col == col:
            try:
                ws.unmerge_cells(str(rng))
            except KeyError:
                # openpyxl can retain a stale merged-cell entry after insert_cols;
                # remove the range metadata so the newly-created column can be rebuilt.
                try:
                    ws.merged_cells.ranges.remove(rng)
                except ValueError:
                    pass


def copy_column_with_vertical_merges(ws, src_col, dst_col, clear_destination=True):
    """Copy a single column's values/styles and vertical merge structure."""
    copy_col_dimension(ws, src_col, dst_col)
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, src_col)
        dst = ws.cell(r, dst_col)
        if not isinstance(dst, MergedCell):
            copy_cell_style(src, dst)
            if clear_destination and r != 1:
                dst.value = None

    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == src_col and rng.max_col == src_col:
            ws.merge_cells(start_row=rng.min_row, start_column=dst_col,
                           end_row=rng.max_row, end_column=dst_col)
            top = ws.cell(rng.min_row, dst_col)
            copy_cell_style(ws.cell(rng.min_row, src_col), top)
            if clear_destination:
                top.value = None


def copy_vertical_merge_structure(ws, src_col, dst_col, header_row, clear=True):
    copy_col_dimension(ws, src_col, dst_col)
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, src_col)
        dst = ws.cell(r, dst_col)
        copy_cell_style(src, dst)
        if clear and r != header_row and not isinstance(dst, MergedCell):
            dst.value = None
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == src_col and rng.max_col == src_col:
            ws.merge_cells(start_row=rng.min_row, start_column=dst_col,
                           end_row=rng.max_row, end_column=dst_col)
            top = ws.cell(rng.min_row, dst_col)
            top.value = None
            copy_cell_style(ws.cell(rng.min_row, src_col), top)


def merge_contiguous_date_time(ws, date_col, time_col, header_row):
    """Fill and merge contiguous Date and Time blocks.

    The uploaded Excel often stores a merged Time cell with the value only in
    the top row.  After unmerging, the lower rows become blank.  We therefore
    carry the current time down within each Date block before merging again.
    This is essential for Schedule for Squad and Session Count calculations.
    """
    if not date_col or not time_col:
        return

    unmerge_ranges_in_column(ws, date_col)
    unmerge_ranges_in_column(ws, time_col)

    # Fill Date downward.
    current_date = None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        if v is not None and str(v).strip() != "":
            current_date = v
        elif current_date is not None:
            ws.cell(r, date_col).value = current_date

    # Fill Time downward, resetting when a new explicit time appears.
    # The reset at a new Date keeps a time from leaking into the next date.
    current_date_key = None
    current_time = None
    for r in range(header_row + 1, ws.max_row + 1):
        dk = date_key(ws.cell(r, date_col).value)
        if dk != current_date_key:
            current_date_key = dk
            current_time = None
        v = ws.cell(r, time_col).value
        if v is not None and str(v).strip() != "":
            current_time = v
        elif current_time is not None:
            ws.cell(r, time_col).value = current_time

    # Merge contiguous equal Dates.
    rows = list(range(header_row + 1, ws.max_row + 1))
    i = 0
    while i < len(rows):
        r0 = rows[i]
        v0 = date_key(ws.cell(r0, date_col).value)
        if not v0:
            i += 1
            continue
        j = i + 1
        while j < len(rows) and date_key(ws.cell(rows[j], date_col).value) == v0:
            j += 1
        r1 = rows[j - 1]
        if r1 > r0:
            ws.merge_cells(start_row=r0, start_column=date_col,
                           end_row=r1, end_column=date_col)
        i = j

    # Find Date blocks from the newly-created merges.
    date_blocks = []
    i = header_row + 1
    while i <= ws.max_row:
        dk = date_key(ws.cell(i, date_col).value)
        if not dk:
            i += 1
            continue
        end = i
        for rng in ws.merged_cells.ranges:
            if (rng.min_col == date_col and rng.max_col == date_col
                    and rng.min_row == i):
                end = rng.max_row
                break
        date_blocks.append((i, end))
        i = end + 1

    # Merge contiguous equal Times inside each Date block.
    for d0, d1 in date_blocks:
        i = d0
        while i <= d1:
            tv = norm_time(ws.cell(i, time_col).value)
            if not tv:
                i += 1
                continue
            j = i + 1
            while j <= d1 and norm_time(ws.cell(j, time_col).value) == tv:
                j += 1
            if j - 1 > i:
                ws.merge_cells(start_row=i, start_column=time_col,
                               end_row=j - 1, end_column=time_col)
            i = j



def set_alignment(ws, col, horizontal, vertical="center", header_row=None):
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(r, col)
        if isinstance(cell, MergedCell):
            continue
        cell.alignment = copy(cell.alignment)
        cell.alignment = Alignment(
            horizontal=horizontal,
            vertical=vertical,
            text_rotation=cell.alignment.text_rotation,
            wrap_text=cell.alignment.wrap_text,
            shrink_to_fit=cell.alignment.shrink_to_fit,
            indent=cell.alignment.indent,
        )


def ensure_target_columns(ws, header_row):
    existing = {norm_header(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1)}
    found = [existing.get(norm_header(h)) for h in TARGET_HEADERS]
    if all(found):
        return dict(zip(TARGET_HEADERS, found))

    start = ws.max_column + 1
    date_col = find_col(ws, {"date"}, header_row)
    time_col = find_col(ws, {"time"}, header_row)
    if not date_col or not time_col:
        raise ValueError("Date and Time columns are required for DPR Formatting.")

    cols = {}
    for idx, h in enumerate(TARGET_HEADERS):
        c = start + idx
        src_col = date_col if h == "Sessions" else time_col
        copy_vertical_merge_structure(ws, src_col, c, header_row, clear=True)
        ws.cell(header_row, c).value = h
        cols[h] = c
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            cell.font = f
    return cols


def process_formatting(wb):
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    time_col = find_col(ws, {"time"}, header_row)
    paper_code_col = find_col(ws, {"paper code"}, header_row)
    missing = []
    if not date_col:
        missing.append("Date")
    if not time_col:
        missing.append("Time")
    if not paper_code_col:
        missing.append("Paper Code")
    if missing:
        raise ValueError("Required headings missing: " + ", ".join(missing))

    merge_contiguous_date_time(ws, date_col, time_col, header_row)

    # Remove exactly two blank spacer columns immediately after Stu. No.
    # This matches the source layout shown in the user's workbook.
    student_col = student_col_for(ws, header_row)
    if student_col + 2 <= ws.max_column:
        if (norm_header(ws.cell(header_row, student_col + 1).value) == "" and
                norm_header(ws.cell(header_row, student_col + 2).value) == ""):
            ws.delete_cols(student_col + 1, 2)

    cols = ensure_target_columns(ws, header_row)

    # Rename the four calculated columns at the DPR stage too, so the
    # downloaded file already has the requested final headings.
    ws.cell(header_row, cols["Total"]).value = "Total Students"
    ws.cell(header_row, cols["Block"]).value = "No. of Blocks"
    ws.cell(header_row, cols["Super"]).value = "Supervisors"
    ws.cell(header_row, cols["Squad"]).value = "Int. Squad"

    # Put the two supervisor columns immediately after Int. Squad and before Sessions.
    existing_int = find_col(ws, {"int. sr. supervisor"}, header_row)
    existing_ext = find_col(ws, {"ext. sr. supervisor"}, header_row)
    if not existing_int and not existing_ext:
        squad_col = find_col(ws, {"int. squad"}, header_row)
        sessions_col = find_col(ws, {"sessions"}, header_row)
        if squad_col and sessions_col:
            ws.insert_cols(squad_col + 1, amount=2)
            sessions_col += 2
            date_col = find_col(ws, {"date"}, header_row) or date_col
            copy_date_format_and_merges(ws, date_col, squad_col + 1, header_row, 1)
            copy_date_format_and_merges(ws, date_col, squad_col + 2, header_row, 2)
            for c, value in ((squad_col + 1, 1), (squad_col + 2, 2)):
                for r in range(header_row + 1, ws.max_row + 1):
                    cell = ws.cell(r, c)
                    if isinstance(cell, MergedCell):
                        continue
                    cell.value = value
                    cell.number_format = "General"
            ws.cell(header_row, squad_col + 1).value = "Int. Sr. Supervisor"
            ws.cell(header_row, squad_col + 2).value = "Ext. Sr. Supervisor"
            ws.cell(header_row, sessions_col).value = "Sessions"
            for r in range(header_row + 1, ws.max_row + 1):
                cell = ws.cell(r, sessions_col)
                if not isinstance(cell, MergedCell):
                    cell.value = None

    # Rebuild the column map after any insertion.
    cols = {}
    for name in ("Total Students", "No. of Blocks", "Supervisors", "Int. Squad", "Int. Sr. Supervisor", "Ext. Sr. Supervisor", "Sessions"):
        c = find_col(ws, {name}, header_row)
        if c:
            cols[name] = c

    # Final requested alignment for original Date/Time columns.
    set_alignment(ws, date_col, "left", "center")
    set_alignment(ws, time_col, "right", "center")

    # Final numeric columns are centered, bold, and slightly larger.
    for h in ("Total Students", "No. of Blocks", "Supervisors", "Int. Squad", "Int. Sr. Supervisor", "Ext. Sr. Supervisor"):
        c = cols.get(h)
        if not c:
            continue
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            f.sz = max((f.sz or 11) + 1, 12)
            cell.font = f

    sessions_c = cols.get("Sessions")
    if sessions_c:
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, sessions_c)
            if not isinstance(cell, MergedCell):
                cell.value = None
                cell.alignment = Alignment(horizontal="center", vertical="center")

    wrap_all_headers(ws, header_row)
    return wb


def student_col_for(ws, header_row):
    # Prefer a recognizable Student/Stu. No. heading when available.
    # If the heading was renamed, the source workbook convention is that
    # the Student Number column is the last source-data column.  Generated
    # DPR/D Form columns are skipped so this remains true after processing.
    c = find_col(ws, {"student", "stu.no", "stu no", "student no", "student number"}, header_row)
    if c:
        return c

    generated_headers = {
        "total", "total students", "block", "blocks", "blocks--", "no. of blocks",
        "super", "super--", "supervisors", "squad", "squad--", "int. squad",
        "int. sr. supervisor", "ext. sr. supervisor", "sessions"
    }
    for c in range(ws.max_column, 0, -1):
        if norm_header(ws.cell(header_row, c).value) not in generated_headers:
            return c

    raise ValueError("Could not find the Student Number column (expected it to be the last source-data column).")


def wrap_all_headers(ws, header_row):
    """Wrap every header cell so long headings do not force wide columns."""
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(header_row, c)
        if isinstance(cell, MergedCell):
            continue
        old = copy(cell.alignment)
        cell.alignment = Alignment(
            horizontal=old.horizontal or "center",
            vertical="center",
            text_rotation=old.text_rotation,
            wrap_text=True,
            shrink_to_fit=old.shrink_to_fit,
            indent=old.indent,
        )

    ws.row_dimensions[header_row].height = max(ws.row_dimensions[header_row].height or 15, 30)


def numeric(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        try:
            return float(s)
        except Exception:
            return None
    return None



def copy_date_format_and_merges(ws, date_col, dst_col, header_row, value):
    """Copy Date-column formatting/vertical merge structure to a new D Form column.
    Populate each visible (unmerged or merged top-left) cell with the requested value.
    """
    # Destination columns are newly inserted/blank; do not unmerge stale ranges
    # left behind by openpyxl after column insertion.
    copy_col_dimension(ws, date_col, dst_col)

    # Copy the Date column's cell formatting exactly.
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, date_col)
        dst = ws.cell(r, dst_col)
        if not isinstance(dst, MergedCell):
            copy_cell_style(src, dst)

    # Recreate the Date column's vertical merged structure.
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == date_col and rng.max_col == date_col:
            ws.merge_cells(start_row=rng.min_row, start_column=dst_col,
                           end_row=rng.max_row, end_column=dst_col)
            copy_cell_style(ws.cell(rng.min_row, date_col), ws.cell(rng.min_row, dst_col))

    # Header keeps its own heading; all data/visible cells receive the requested value.
    ws.cell(header_row, dst_col).alignment = copy(ws.cell(header_row, date_col).alignment)
    ws.cell(header_row, dst_col).value = None
    for r in range(header_row + 1, ws.max_row + 1):
        cell = ws.cell(r, dst_col)
        if isinstance(cell, MergedCell):
            continue
        cell.value = value


def move_sessions_to_final_position(ws, header_row, sessions_col):
    """Make final D Form tail contiguous: ... Int. Squad | Int. Sr. Superior | Ext. Sr. Supervisor | Sessions.
    The original Sessions column is copied to the new final position, preserving its vertical merge structure.
    """
    # Reserve two contiguous new columns immediately after the old Sessions column.
    # Final layout is: Int. Sr. Superior | Ext. Sr. Supervisor | Sessions.
    final_sessions_col = max(ws.max_column + 2, sessions_col + 2)
    # Copy original Sessions structure to the new final column before clearing the source.
    unmerge_ranges_in_column(ws, final_sessions_col)
    copy_col_dimension(ws, sessions_col, final_sessions_col)
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, sessions_col)
        dst = ws.cell(r, final_sessions_col)
        if not isinstance(dst, MergedCell):
            copy_cell_style(src, dst)
            dst.value = src.value
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == sessions_col and rng.max_col == sessions_col:
            ws.merge_cells(start_row=rng.min_row, start_column=final_sessions_col,
                           end_row=rng.max_row, end_column=final_sessions_col)
            top = ws.cell(rng.min_row, final_sessions_col)
            src_top = ws.cell(rng.min_row, sessions_col)
            copy_cell_style(src_top, top)
            top.value = src_top.value

    # Remove the old Sessions merge structure, then use it as Int. Sr. Superior.
    unmerge_ranges_in_column(ws, sessions_col)
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(r, sessions_col)
        if not isinstance(cell, MergedCell):
            cell.value = None
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font); f.bold = True; cell.font = f

    # The new Ext. Sr. Supervisor column is inserted logically by using the next column.
    ext_col = sessions_col + 1
    if ext_col != final_sessions_col:
        # There should be no unrelated column between the old Sessions and appended final Sessions.
        # Copy any current content/styles only if necessary, then clear it for the new column.
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(r, ext_col)
            if not isinstance(cell, MergedCell):
                cell.value = None
                cell.alignment = Alignment(horizontal="center", vertical="center")
        copy_col_dimension(ws, sessions_col, ext_col)

    ws.cell(header_row, sessions_col).value = "Int. Sr. Superior"
    ws.cell(header_row, ext_col).value = "Ext. Sr. Supervisor"
    ws.cell(header_row, final_sessions_col).value = "Sessions"

    for c in (sessions_col, ext_col, final_sessions_col):
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font); f.bold = True; cell.font = f
            if c == final_sessions_col:
                cell.value = None

    return final_sessions_col


def process_dform(wb):
    ws = get_main_sheet(wb)
    header_row = (find_header_row(ws, {"total students", "no. of blocks", "supervisors", "int. squad"})
                  or find_header_row(ws, {"total", "block", "super", "squad"})
                  or 1)
    student_col = student_col_for(ws, header_row)

    # Remove exactly the two blank spacer columns immediately after Stu. No.
    # Do this only when they are actually blank, so real data columns are never deleted.
    if student_col + 2 <= ws.max_column:
        gap1 = ws.cell(header_row, student_col + 1).value
        gap2 = ws.cell(header_row, student_col + 2).value
        if norm_header(gap1) == "" and norm_header(gap2) == "":
            ws.delete_cols(student_col + 1, 2)

    # Find the calculation columns, accepting the old headings as well.
    def find_any_col(*names):
        wanted = {norm_header(x) for x in names}
        for c in range(1, ws.max_column + 1):
            if norm_header(ws.cell(header_row, c).value) in wanted:
                return c
        return None

    total_col = find_any_col("Total", "Total Students")
    block_col = find_any_col("Block", "Blocks", "Blocks--", "No. of Blocks")
    super_col = find_any_col("Super", "Super--", "Supervisors")
    squad_col = find_any_col("Squad", "Squad--", "Int. Squad")
    sessions_col = find_any_col("Sessions")
    if not all([total_col, block_col, super_col, squad_col, sessions_col]):
        raise ValueError("Required Total/Block/Super/Squad/Sessions columns were not found.")

    # Calculate one Total/Block/Super/Squad set for each Date+Time session.
    # A merged Time cell represents one session even when several class rows
    # belong to it.  Sum all student counts in that session, then calculate
    # blocks/supervisors/squad from the session total and merge the result
    # columns across the same rows.
    date_col_for_group = find_col(ws, {"date"}, header_row)
    time_col_for_group = find_col(ws, {"time"}, header_row)
    if date_col_for_group and time_col_for_group:
        calc_cols = [total_col, block_col, super_col, squad_col]
        for c in calc_cols:
            unmerge_ranges_in_column(ws, c)

        # Build contiguous Date+Time groups using inherited values.
        # Merged Excel cells store the value only in the top row; lower rows
        # are MergedCell/blank. Carry the last Date and Time forward so the
        # whole merged session is treated as ONE calculation group.
        groups = []
        current_date_key = None
        current_time_key = None
        start = None
        for r in range(header_row + 1, ws.max_row + 1):
            raw_d = ws.cell(r, date_col_for_group).value
            raw_t = ws.cell(r, time_col_for_group).value
            dk = date_key(raw_d) if raw_d not in (None, "") else current_date_key
            tv = norm_time(raw_t) if raw_t not in (None, "") else current_time_key

            # A new explicit date resets the inherited time.
            explicit_d = date_key(raw_d) if raw_d not in (None, "") else None
            if explicit_d is not None and explicit_d != current_date_key:
                current_date_key = explicit_d
                current_time_key = None
                tv = norm_time(raw_t) if raw_t not in (None, "") else None
            else:
                current_date_key = dk

            if raw_t not in (None, ""):
                current_time_key = norm_time(raw_t)
                tv = current_time_key
            else:
                tv = current_time_key

            if not current_date_key or not tv:
                if start is not None:
                    groups.append((start, r - 1))
                    start = None
                continue

            key = (current_date_key, tv)
            if start is None:
                start = r
                prev_key = key
            elif key != prev_key:
                groups.append((start, r - 1))
                start = r
                prev_key = key

        if start is not None:
            groups.append((start, ws.max_row))

        for g0, g1 in groups:
            total = 0
            found = False
            for rr in range(g0, g1 + 1):
                n = numeric(ws.cell(rr, student_col).value)
                if n is not None:
                    total += n
                    found = True
            total = total if found else 0

            if 0 <= total <= 35:
                block = 1
            elif 36 <= total <= 70:
                block = 2
            elif 71 <= total <= 105:
                block = 3
            elif 106 <= total <= 135:
                block = 4
            elif 136 <= total <= 165:
                block = 5
            elif 166 <= total <= 195:
                block = 6
            elif 196 <= total <= 225:
                block = 7
            else:
                block = None

            squad = None if block is None else (0 if block <= 2 else (2 if block <= 4 else 3))

            values = {
                total_col: total,
                block_col: block,
                super_col: block,
                squad_col: squad,
            }
            for c, value in values.items():
                top = ws.cell(g0, c)
                if not isinstance(top, MergedCell):
                    top.value = value
                if g1 > g0:
                    ws.merge_cells(start_row=g0, start_column=c,
                                   end_row=g1, end_column=c)

    # Calculate Total / Block / Super / Squad.
    merged_total_rows = set()
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == total_col and rng.max_col == total_col and rng.min_row > header_row:
            merged_total_rows.update(range(rng.min_row, rng.max_row + 1))
            total = 0
            found = False
            for r in range(rng.min_row, rng.max_row + 1):
                n = numeric(ws.cell(r, student_col).value)
                if n is not None:
                    total += n
                    found = True
            top = ws.cell(rng.min_row, total_col)
            if not isinstance(top, MergedCell):
                top.value = total if found else 0

    for r in range(header_row + 1, ws.max_row + 1):
        tc = ws.cell(r, total_col)
        if r not in merged_total_rows and not isinstance(tc, MergedCell):
            tc.value = ws.cell(r, student_col).value

    # Fill only rows that are not already part of a merged calculation group.
    for r in range(header_row + 1, ws.max_row + 1):
        if any(rng.min_row <= r <= rng.max_row and rng.min_col == total_col and rng.max_col == total_col
               for rng in ws.merged_cells.ranges):
            continue
        total = numeric(ws.cell(r, total_col).value)
        if total is None:
            continue
        if 0 <= total <= 35:
            block = 1
        elif 36 <= total <= 70:
            block = 2
        elif 71 <= total <= 105:
            block = 3
        elif 106 <= total <= 135:
            block = 4
        elif 136 <= total <= 165:
            block = 5
        elif 166 <= total <= 195:
            block = 6
        elif 196 <= total <= 225:
            block = 7
        else:
            continue
        squad = 0 if block <= 2 else (2 if block <= 4 else 3)
        for c, value in ((block_col, block), (super_col, block), (squad_col, squad)):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell.value = value

    # Exact requested headings.
    ws.cell(header_row, total_col).value = "Total Students"
    ws.cell(header_row, block_col).value = "No. of Blocks"
    ws.cell(header_row, super_col).value = "Supervisors"
    ws.cell(header_row, squad_col).value = "Int. Squad"

    # Remove any previously-created supervisor columns if this D Form step is run
    # repeatedly on an already-processed workbook. Keep the Sessions column.
    old_supervisor_cols = []
    for c in range(1, ws.max_column + 1):
        h = norm_header(ws.cell(header_row, c).value)
        if h in {
            "int. sr. supervisor", "int. sr. superior",
            "ext. sr. supervisor", "ext. sr. superior"
        }:
            old_supervisor_cols.append(c)
    for c in reversed(old_supervisor_cols):
        if c != sessions_col:
            ws.delete_cols(c, 1)
            if c < sessions_col:
                sessions_col -= 1

    # Refresh Squad/Sessions positions after any cleanup.
    squad_col = find_any_col("Int. Squad", "Squad", "Squad--")
    sessions_col = find_any_col("Sessions")
    if not squad_col or not sessions_col:
        raise ValueError("Int. Squad or Sessions column was lost while preparing D Form.")

    # Insert the two requested columns immediately after Int. Squad.
    int_sr_col = squad_col + 1
    ws.insert_cols(int_sr_col, amount=2)
    ext_sr_col = int_sr_col + 1

    # Copy Date-column formatting and vertical merge structure to both columns.
    # This also clears the old date values in the new columns.
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        raise ValueError("Date column is required for supervisor columns.")
    copy_date_format_and_merges(ws, date_col, int_sr_col, header_row, 1)
    copy_date_format_and_merges(ws, date_col, ext_sr_col, header_row, 2)

    # The Date column's number format would otherwise display 1/2 as dates
    # (for example 01/01/1900 and 02/01/1900). Keep the Date-style layout,
    # but make these supervisor values real numeric 1/2 with General format.
    for c, value in ((int_sr_col, 1), (ext_sr_col, 2)):
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.value = value
            cell.number_format = "General"

    # Sessions must use the same formatting/merge structure as Ext. Sr. Supervisor,
    # while remaining completely blank. Rebuild the Sessions column from Ext. Sr. Supervisor.
    sessions_col = find_any_col("Sessions")
    unmerge_ranges_in_column(ws, sessions_col)
    copy_col_dimension(ws, ext_sr_col, sessions_col)
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, ext_sr_col)
        dst = ws.cell(r, sessions_col)
        if not isinstance(dst, MergedCell):
            copy_cell_style(src, dst)
            dst.value = None
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == ext_sr_col and rng.max_col == ext_sr_col:
            ws.merge_cells(start_row=rng.min_row, start_column=sessions_col,
                           end_row=rng.max_row, end_column=sessions_col)
            top = ws.cell(rng.min_row, sessions_col)
            top.value = None
            copy_cell_style(ws.cell(rng.min_row, ext_sr_col), top)

    ws.cell(header_row, int_sr_col).value = "Int. Sr. Supervisor"
    ws.cell(header_row, ext_sr_col).value = "Ext. Sr. Supervisor"
    ws.cell(header_row, sessions_col).value = "Sessions"

    # Center, bold, and slightly enlarge all numeric values from Total Students
    # through Ext. Sr. Supervisor.
    numeric_format_cols = (
        total_col, block_col, super_col, squad_col, int_sr_col, ext_sr_col
    )
    for c in numeric_format_cols:
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            base_size = f.sz if f.sz else 11
            f.sz = max(base_size + 1, 12)
            cell.font = f

    return wb


def session_count(times):
    u = {norm_time(t) for t in times if norm_time(t)}
    pairs = [
        {"9-10:30 am", "9-11 am"},
        {"12-1:30 pm", "12-2 pm"},
        {"3-4:30 pm", "3-5 pm"},
    ]
    used = set()
    count = 0
    for pair in pairs:
        if pair.issubset(u):
            count += 1
            used.update(pair)
    return count + len(u - used)


def apply_summary_borders(ws):
    side = Side(style="thin")
    border = Border(left=side, right=side, top=side, bottom=side)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    # The Summary "Dates" column must be left aligned.
    if ws.max_column >= 2:
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(r, 2)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def extract_workbook_dates(ws, header_row=1):
    """Return sorted unique actual dates present in the main sheet."""
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        return []
    out = set()
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        k = date_key(v)
        if k:
            try:
                out.add(datetime.strptime(k, "%Y-%m-%d").date())
            except ValueError:
                pass
    return sorted(out)


def add_date_range_report(wb, start_date, end_date):
    """Filter the main worksheet in-place to the selected inclusive date range.
    Preserves cell values, styles, dimensions and applicable merged ranges.
    """
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        raise ValueError("Date column सापडला नाही.")

    row_dates = {}
    current = None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        if v not in (None, ""):
            k = date_key(v)
            try:
                current = datetime.strptime(k, "%Y-%m-%d").date()
            except (ValueError, TypeError):
                current = None
        row_dates[r] = current

    original_merges = [str(rng) for rng in ws.merged_cells.ranges]
    keep_rows = {
        r for r in range(header_row + 1, ws.max_row + 1)
        if row_dates.get(r) is not None and start_date <= row_dates[r] <= end_date
    }

    # Unmerge before deleting rows so merged-cell placeholders do not block deletion.
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    old_max_row = ws.max_row
    deleted_before = 0
    row_map = {}
    for r in range(1, old_max_row + 1):
        if r <= header_row or r in keep_rows:
            row_map[r] = r - deleted_before
        else:
            deleted_before += 1

    for r in range(old_max_row, header_row, -1):
        if r not in keep_rows:
            ws.delete_rows(r, 1)

    # Recreate merged ranges only when the entire original range survived filtering.
    for merge_ref in original_merges:
        min_col, min_row, max_col, max_row = range_boundaries(merge_ref)
        if all(rr <= header_row or rr in keep_rows for rr in range(min_row, max_row + 1)):
            new_min_row = row_map[min_row]
            new_max_row = row_map[max_row]
            try:
                ws.merge_cells(start_row=new_min_row, start_column=min_col,
                               end_row=new_max_row, end_column=max_col)
            except Exception:
                pass

    return ws

def build_date_range_report(wb, start_date, end_date):
    """Return a workbook whose main report sheet contains only the selected date range.
    If either bound is None, the full available date range is used.
    """
    src = get_main_sheet(wb)
    original_title = src.title
    all_dates = extract_workbook_dates(src, find_header_row(src, {"date", "time"}) or 1)
    if not all_dates:
        raise ValueError("Date column मध्ये कोणतीही तारीख सापडली नाही.")
    if start_date is None:
        start_date = all_dates[0]
    if end_date is None:
        end_date = all_dates[-1]
    if start_date > end_date:
        raise ValueError("From Date हा To Date पेक्षा मोठा असू शकत नाही.")

    # Keep a copy of session counts before replacing the main sheet.
    session_counts = {}
    header_row = find_header_row(src, {"date", "time"}) or 1
    date_col = find_col(src, {"date"}, header_row)
    sessions_col = find_col(src, {"sessions"}, header_row)
    current_date = None
    for r in range(header_row + 1, src.max_row + 1):
        dv = src.cell(r, date_col).value
        if dv not in (None, ""):
            k = date_key(dv)
            try:
                current_date = datetime.strptime(k, "%Y-%m-%d").date()
            except Exception:
                current_date = None
        if current_date and sessions_col:
            sv = src.cell(r, sessions_col).value
            if sv not in (None, ""):
                try:
                    session_counts[current_date] = int(sv)
                except Exception:
                    pass

    filtered = add_date_range_report(wb, start_date, end_date)
    # Keep the main worksheet visible; filtering it in-place avoids an empty-workbook error.
    filtered.title = original_title

    # For Session Count reports, also filter Session Summary to the same date range.
    if "Session Summary" in wb.sheetnames and session_counts:
        del wb["Session Summary"]
        summary = wb.create_sheet("Session Summary")
        headers = ["Session Name", "Dates", "Total Dates", "Actual Sessions", "Rs."]
        summary.append(headers)
        for c in range(1, 6):
            summary.cell(1, c).font = Font(bold=True)
            summary.cell(1, c).alignment = Alignment(horizontal="center", vertical="center")
        cats = {i: [] for i in range(1, 8)}
        selected_counts = {d:n for d,n in session_counts.items() if start_date <= d <= end_date}
        for d, n in selected_counts.items():
            if 1 <= n <= 7:
                cats[n].append(d.strftime("%Y-%m-%d"))
        for n in range(1, 8):
            ds = sorted(cats[n])
            total = len(ds)
            actual = n * total
            summary.append([n, ", ".join(display_date(d) for d in ds), total, actual, actual * 105])
        total_dates = len(selected_counts)
        total_sessions = sum(selected_counts.values())
        summary.append(["TOTAL", "", total_dates, total_sessions, total_sessions * 105])
        summary.column_dimensions["A"].width = 15
        summary.column_dimensions["B"].width = 55
        summary.column_dimensions["C"].width = 15
        summary.column_dimensions["D"].width = 18
        summary.column_dimensions["E"].width = 15
        apply_summary_borders(summary)

    return wb


def resolve_date_range(start_choice, end_choice, available_dates):
    """Convert dropdown selections to an actual date range; ALL means full range."""
    if not available_dates:
        raise ValueError("Date column मध्ये कोणतीही तारीख सापडली नाही.")
    if start_choice == "ALL" or end_choice == "ALL":
        return available_dates[0], available_dates[-1]
    return start_choice, end_choice


def process_sessions(wb):
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    time_col = find_col(ws, {"time"}, header_row)
    if not date_col or not time_col:
        raise ValueError("Date and Time columns are required for Session Count.")

    by_date = {}
    current_date = ""
    for r in range(header_row + 1, ws.max_row + 1):
        d = ws.cell(r, date_col).value
        if d is not None and str(d).strip() != "":
            current_date = date_key(d)
        t = ws.cell(r, time_col).value
        if current_date and t is not None and str(t).strip() != "":
            by_date.setdefault(current_date, []).append(t)

    counts = {d: session_count(ts) for d, ts in by_date.items() if d}

    sessions_col = find_col(ws, {"sessions"}, header_row)
    if sessions_col:
        # Fill the Sessions column with the calculated session count for each date.
        # The Sessions column has the same vertical merge structure as Ext. Sr. Supervisor,
        # so the value is written to the visible/top-left cell of each merged date group.
        current_date = ""
        date_for_row = {}
        for r in range(header_row + 1, ws.max_row + 1):
            d = ws.cell(r, date_col).value
            if d is not None and str(d).strip() != "":
                current_date = date_key(d)
            date_for_row[r] = current_date

        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, sessions_col)
            if isinstance(cell, MergedCell):
                continue
            d = date_for_row.get(r, "")
            cell.value = counts.get(d, None)
            cell.number_format = "General"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            f.sz = max((f.sz or 11) + 1, 12)
            cell.font = f

    if "Session Summary" in wb.sheetnames:
        del wb["Session Summary"]
    summary = wb.create_sheet("Session Summary")
    headers = ["Session Name", "Dates", "Total Dates", "Actual Sessions", "Rs."]
    summary.append(headers)
    for c in range(1, 6):
        summary.cell(1, c).font = Font(bold=True)
        summary.cell(1, c).alignment = Alignment(horizontal="center", vertical="center")

    cats = {i: [] for i in range(1, 8)}
    for d, n in counts.items():
        if 1 <= n <= 7:
            cats[n].append(d)

    for n in range(1, 8):
        ds = sorted(cats[n])
        total_dates = len(ds)
        actual = n * total_dates
        summary.append([n, ", ".join(display_date(d) for d in ds), total_dates, actual, actual * 105])

    total_dates = len(counts)
    total_sessions = sum(counts.values())
    summary.append(["TOTAL", "", total_dates, total_sessions, total_sessions * 105])

    independent = sum(session_count(ts) for ts in by_date.values())
    if independent != total_sessions:
        raise ValueError(f"Session verification failed: summary={total_sessions}, independent={independent}")

    summary.column_dimensions["A"].width = 15
    summary.column_dimensions["B"].width = 55
    summary.column_dimensions["C"].width = 15
    summary.column_dimensions["D"].width = 18
    summary.column_dimensions["E"].width = 15
    for r in range(2, summary.max_row + 1):
        summary.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    apply_summary_borders(summary)

    return wb, total_dates, total_sessions



def report_source_bytes():
    """Return the full processed workbook, never a date/class-filtered report copy."""
    return (
        st.session_state.get("session_full_bytes")
        or st.session_state.get("dform_bytes")
    )


def workbook_rows_for_range(wb, start_date, end_date):
    """Extract the main sheet rows for an inclusive date range, forward-filling merged dates."""
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        raise ValueError("Date column सापडला नाही.")
    headers = [ws.cell(header_row, c).value or f"Column {c}" for c in range(1, ws.max_column + 1)]
    rows = []
    current_date = None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        if v not in (None, ""):
            k = date_key(v)
            try:
                current_date = datetime.strptime(k, "%Y-%m-%d").date()
            except Exception:
                current_date = None
        if current_date is None or not (start_date <= current_date <= end_date):
            continue
        vals = []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            vals.append(None if isinstance(cell, MergedCell) else cell.value)
        vals[date_col - 1] = current_date
        rows.append(vals)
    return headers, rows


def filtered_workbook_copy(source_bytes, start_date, end_date):
    wb = load_workbook(BytesIO(source_bytes), data_only=False)
    build_date_range_report(wb, start_date, end_date)
    return wb


def make_pdf_report(source_bytes, start_date, end_date, title="Print Report"):
    """Create a print-friendly class/date grouped PDF report.

    Layout:
    - Wide landscape page (no A4 restriction)
    - Each Class starts on a new page
    - Date and Class cells are vertically merged for repeated rows
    - Table header repeats on page breaks
    - Readable 8pt body text with wrapped headers/cells
    """
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    )

    wb = load_workbook(BytesIO(source_bytes), data_only=False)
    headers, rows = workbook_rows_for_range(wb, start_date, end_date)
    out = BytesIO()

    ncols = max(1, len(headers))
    col_width = 25 * mm
    page_width = max(320 * mm, min(650 * mm, ncols * col_width + 20 * mm))
    page_height = 210 * mm

    def draw_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(page_width - 8 * mm, 4 * mm, f"Page {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        out,
        pagesize=(page_width, page_height),
        rightMargin=8 * mm,
        leftMargin=8 * mm,
        topMargin=8 * mm,
        bottomMargin=9 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "PrintTitle", parent=styles["Title"], fontSize=15, leading=18,
        spaceAfter=2 * mm
    )
    normal_style = ParagraphStyle(
        "PrintNormal", parent=styles["Normal"], fontSize=9, leading=11
    )
    section_style = ParagraphStyle(
        "ClassHeading", parent=styles["Heading2"], fontSize=13, leading=15,
        spaceBefore=1 * mm, spaceAfter=3 * mm
    )
    cell_style = ParagraphStyle(
        "PrintCell", parent=styles["Normal"], fontSize=8, leading=9,
        wordWrap="CJK", alignment=0
    )
    header_style = ParagraphStyle(
        "PrintHeader", parent=cell_style, fontName="Helvetica-Bold",
        alignment=1
    )

    def cell_text(v):
        if isinstance(v, datetime):
            return v.strftime("%d/%m/%Y %I:%M %p")
        if isinstance(v, date):
            return v.strftime("%d/%m/%Y")
        return "" if v is None else str(v)

    elements = [
        Paragraph(title, title_style),
        Paragraph(
            f"Date Range: {start_date.strftime('%d/%m/%Y')} to "
            f"{end_date.strftime('%d/%m/%Y')}",
            normal_style,
        ),
        Spacer(1, 3 * mm),
    ]

    if not rows:
        elements.append(Paragraph("No data found for the selected date range.", normal_style))
        doc.build(elements, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
        return out.getvalue()

    # Identify Date/Class columns so their repeated values can be merged visually.
    date_idx = next((i for i, h in enumerate(headers) if norm_header(h) == "date"), None)
    class_idx = None
    class_aliases = {
        "class", "class name", "class_name", "standard", "std", "std.",
        "standard name", "class/standard", "class/div", "class / div",
        "class division", "इयत्ता"
    }
    for i, h in enumerate(headers):
        if norm_header(h) in class_aliases:
            class_idx = i
            break

    # Forward-fill Date/Class for grouping, without changing the underlying report.
    enriched = []
    current_date = None
    current_class = None
    for row in rows:
        vals = list(row)
        if date_idx is not None and vals[date_idx] not in (None, ""):
            current_date = vals[date_idx]
        if class_idx is not None and vals[class_idx] not in (None, ""):
            current_class = str(vals[class_idx]).strip()
        enriched.append((current_date, current_class, vals))

    # Group first by Class, then by Date. If no Class column exists, keep one group.
    groups = []
    grouped = {}
    for dval, cval, vals in enriched:
        class_key = cval or "All Classes"
        grouped.setdefault(class_key, []).append((dval, vals))

    def class_sort_key(item):
        return str(item[0]).lower()

    for class_name, class_rows in sorted(grouped.items(), key=class_sort_key):
        groups.append((class_name, class_rows))

    available_width = page_width - 16 * mm
    widths = [available_width / ncols] * ncols

    for group_no, (class_name, class_rows) in enumerate(groups):
        if group_no > 0:
            elements.append(PageBreak())

        elements.append(Paragraph(f"Class: {cell_text(class_name)}", section_style))
        elements.append(Paragraph(
            f"Records: {len(class_rows)}", normal_style
        ))
        elements.append(Spacer(1, 2 * mm))

        # Build one table per class so Class can span the complete class section.
        data = [[Paragraph(cell_text(h), header_style) for h in headers]]
        for dval, vals in class_rows:
            data.append([Paragraph(cell_text(v), cell_style) for v in vals])

        table = Table(
            data,
            repeatRows=1,
            colWidths=widths,
            hAlign="LEFT",
        )

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]

        # Merge repeated Date cells vertically inside each class table.
        if date_idx is not None and len(class_rows) > 0:
            pos = 1
            while pos <= len(class_rows):
                key = date_key(class_rows[pos - 1][0]) if class_rows[pos - 1][0] not in (None, "") else ""
                end_pos = pos
                while end_pos < len(class_rows):
                    nxt = class_rows[end_pos][0]
                    nxt_key = date_key(nxt) if nxt not in (None, "") else ""
                    if nxt_key != key:
                        break
                    end_pos += 1
                if end_pos - pos > 1 and key:
                    style_cmds.append(("SPAN", (date_idx, pos), (date_idx, end_pos - 1)))
                    style_cmds.append(("VALIGN", (date_idx, pos), (date_idx, end_pos - 1), "MIDDLE"))
                pos = end_pos

        # Merge Class column across the whole table when present.
        if class_idx is not None and len(class_rows) > 1:
            style_cmds.append(("SPAN", (class_idx, 1), (class_idx, len(class_rows))))
            style_cmds.append(("VALIGN", (class_idx, 1), (class_idx, len(class_rows)), "MIDDLE"))

        table.setStyle(TableStyle(style_cmds))
        elements.append(table)

    doc.build(elements, onFirstPage=draw_page_number, onLaterPages=draw_page_number)
    return out.getvalue()


def find_class_col(ws, header_row):
    return find_col(ws, {
        "class", "class name", "class_name", "standard", "std", "std.",
        "standard name", "class/standard", "class/div", "class / div",
        "class division", "इयत्ता"
    }, header_row)


def extract_workbook_classes(ws, header_row=1):
    class_col = find_class_col(ws, header_row)
    if not class_col:
        return []
    values, current = set(), None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, class_col).value
        if v not in (None, ""):
            current = str(v).strip()
        if current:
            values.add(current)
    return sorted(values, key=str.lower)


def available_report_classes(source_bytes):
    wb = load_workbook(BytesIO(source_bytes), data_only=False)
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    return extract_workbook_classes(ws, header_row)


def filter_workbook_by_date_and_class(wb, start_date, end_date, selected_classes=None):
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    class_col = find_class_col(ws, header_row)

    if not date_col:
        raise ValueError("Date column सापडला नाही.")
    if selected_classes and not class_col:
        raise ValueError("Excel मध्ये Class/Standard column सापडला नाही.")

    selected = {str(x).strip() for x in (selected_classes or [])}
    row_info, current_date, current_class = {}, None, None

    for r in range(header_row + 1, ws.max_row + 1):
        dv = ws.cell(r, date_col).value
        if dv not in (None, ""):
            try:
                current_date = datetime.strptime(date_key(dv), "%Y-%m-%d").date()
            except Exception:
                current_date = None
        if class_col:
            cv = ws.cell(r, class_col).value
            if cv not in (None, ""):
                current_class = str(cv).strip()
        row_info[r] = (current_date, current_class)

    keep_rows = {
        r for r, (d, c) in row_info.items()
        if d is not None and start_date <= d <= end_date
        and (not selected or c in selected)
    }

    original_merges = [str(rng) for rng in ws.merged_cells.ranges]
    old_max_row = ws.max_row
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    row_map, deleted_before = {}, 0
    for r in range(1, old_max_row + 1):
        if r <= header_row or r in keep_rows:
            row_map[r] = r - deleted_before
        else:
            deleted_before += 1

    for r in range(old_max_row, header_row, -1):
        if r not in keep_rows:
            ws.delete_rows(r, 1)

    for ref in original_merges:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        if all(rr <= header_row or rr in keep_rows for rr in range(min_row, max_row + 1)):
            try:
                ws.merge_cells(
                    start_row=row_map[min_row], start_column=min_col,
                    end_row=row_map[max_row], end_column=max_col
                )
            except Exception:
                pass
    return wb


def _set_all_checkbox_items(prefix, count):
    """Set all vertical checkbox widget states to match Select All."""
    value = bool(st.session_state.get(f"{prefix}_select_all", True))
    for i in range(count):
        st.session_state[f"{prefix}_item_{i}"] = value


def vertical_checkbox_selector(options, prefix, label="Select", collapsed=True):
    """Reusable vertical checkbox list with a working Select All toggle."""
    options = list(options or [])
    if not options:
        return []

    style_red_checkbox()
    select_key = f"{prefix}_select_all"
    item_keys = [f"{prefix}_item_{i}" for i in range(len(options))]

    if select_key not in st.session_state:
        st.session_state[select_key] = True
    for key in item_keys:
        if key not in st.session_state:
            st.session_state[key] = bool(st.session_state[select_key])

    with st.expander(label, expanded=not collapsed):
        st.checkbox(
            "Select All",
            key=select_key,
            on_change=_set_all_checkbox_items,
            args=(prefix, len(options)),
        )
        selected = []
        for i, option in enumerate(options):
            if st.checkbox(str(option), key=item_keys[i]):
                selected.append(option)

        # Do not write to the Select All widget state after its widget is created.
        # This avoids StreamlitWidgetAlreadyInstantiatedError.

    return selected


def vertical_class_selector(source_bytes, prefix, collapsed=True):
    """Class filter shown collapsed with vertical black labels and red ticks."""
    return vertical_checkbox_selector(
        available_report_classes(source_bytes),
        prefix,
        label="Class",
        collapsed=collapsed,
    )


def date_range_controls(source_bytes, prefix, include_class=False):
    """From/To date range, optionally followed by a collapsed vertical Class filter."""
    dates = available_report_dates(source_bytes)
    if not dates:
        return None, None, []
    start_date = st.selectbox(
        "From Date", dates,
        format_func=lambda d: d.strftime("%d/%m/%Y"),
        key=f"{prefix}_from",
    )
    valid_end_dates = [d for d in dates if d >= start_date]
    default_end = st.session_state.get(f"{prefix}_to", valid_end_dates[-1])
    if default_end not in valid_end_dates:
        default_end = valid_end_dates[-1]
    end_date = st.selectbox(
        "To Date", valid_end_dates,
        index=valid_end_dates.index(default_end),
        format_func=lambda d: d.strftime("%d/%m/%Y"),
        key=f"{prefix}_to",
    )
    selected_classes = vertical_class_selector(source_bytes, prefix, collapsed=True) if include_class else []
    return start_date, end_date, selected_classes


def date_class_controls(source_bytes, prefix):
    """Backward-compatible Date + collapsed Class filter."""
    return date_range_controls(source_bytes, prefix, include_class=True)


def all_date_range(source_bytes):
    dates = available_report_dates(source_bytes)
    if not dates:
        return None, None
    return dates[0], dates[-1]


def apply_excel_print_formatting(wb):
    """Apply print-friendly Excel page setup without changing report data/calculations."""
    for ws in wb.worksheets:
        if ws.max_row < 1 or ws.max_column < 1:
            continue

        header_row = find_header_row(ws, {"date", "time"})
        if header_row is None:
            # Summary/other sheets: use the first row as the header when it contains data.
            header_row = 1

        # Freeze the report header so scrolling/printing is easier.
        ws.freeze_panes = f"A{header_row + 1}"
        ws.sheet_view.showGridLines = False

        # Print settings: wide reports should print landscape and fit to one page wide.
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.45
        ws.page_margins.bottom = 0.45
        ws.page_margins.header = 0.2
        ws.page_margins.footer = 0.2

        # Repeat the header row on every printed page.
        ws.print_title_rows = f"${header_row}:${header_row}"
        ws.print_options.horizontalCentered = False
        ws.sheet_properties.pageSetUpPr.autoPageBreaks = False

        # Define the print area to the actual used range.
        ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

        # Width is based on DATA, not header length. Headers are wrapped, so
        # long headings can occupy multiple lines without making columns huge.
        for c in range(1, ws.max_column + 1):
            letter = get_column_letter(c)
            max_len = 0
            sample_end = min(ws.max_row, header_row + 250)
            for r in range(header_row + 1, sample_end + 1):
                cell = ws.cell(r, c)
                if isinstance(cell, MergedCell):
                    continue
                value = cell.value
                if value is None:
                    continue
                text = str(value)
                if len(text) > max_len:
                    max_len = len(text)
            width = min(max(max_len + 2, 8), 18)
            ws.column_dimensions[letter].width = width

        wrap_all_headers(ws, header_row)

        # Keep date/time values readable and vertically centered.
        for row in ws.iter_rows(min_row=header_row, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                cell.alignment = copy(cell.alignment)
                cell.alignment = Alignment(
                    horizontal=cell.alignment.horizontal or "left",
                    vertical="center",
                    wrap_text=True,
                    text_rotation=cell.alignment.textRotation,
                    shrink_to_fit=cell.alignment.shrinkToFit,
                    indent=cell.alignment.indent,
                )

        # Make the header easy to identify in the printed sheet while preserving
        # the existing workbook's overall style.
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(header_row, c)
            if isinstance(cell, MergedCell):
                continue
            f = copy(cell.font)
            f.bold = True
            cell.font = f
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.row_dimensions[header_row].height = max(ws.row_dimensions[header_row].height or 15, 28)

        # Print footer with page number and sheet name.
        ws.oddFooter.center.text = "Page &[Page] of &[Pages]"
        ws.oddFooter.right.text = "&[Tab]"


def filtered_report_bytes(source, start_date, end_date, selected_classes):
    wb = load_workbook(BytesIO(source), data_only=False)
    filter_workbook_by_date_and_class(wb, start_date, end_date, selected_classes)
    apply_excel_print_formatting(wb)
    return wb_bytes(wb)


def available_report_dates(source_bytes):
    wb = load_workbook(BytesIO(source_bytes), data_only=False)
    ws = get_main_sheet(wb)
    return extract_workbook_dates(ws, find_header_row(ws, {"date", "time"}) or 1)


for key in ("original_bytes", "format_bytes", "dform_bytes", "final_bytes", "range_report_bytes", "dform_range_report_bytes", "report_range", "dform_report_range", "datewise_excel_bytes", "print_excel_bytes", "print_pdf_bytes", "session_full_bytes", "session_report_bytes", "t1_report_excel_bytes", "t1_report_pdf_bytes"):
    if key not in st.session_state:
        st.session_state[key] = None


def render_print_report(source, start_date, end_date, selected_classes, prefix, title="DPR / D Form Print Report"):
    """Direct print/download actions: no separate Generate step."""
    st.markdown("---")
    st.subheader("🖨️ Print Report")
    st.caption(
        f"Print कालावधी: {start_date.strftime('%d/%m/%Y')} ते "
        f"{end_date.strftime('%d/%m/%Y')}"
    )

    try:
        filtered_bytes = filtered_report_bytes(
            source, start_date, end_date, selected_classes
        )
        print_excel = filtered_bytes
        print_pdf = make_pdf_report(
            filtered_bytes, start_date, end_date, title
        )

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "🖨️ Print Excel",
                print_excel,
                f"{prefix}_Print_Report.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{prefix}_direct_print_excel",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "🖨️ Print PDF",
                print_pdf,
                f"{prefix}_Print_Report.pdf",
                "application/pdf",
                key=f"{prefix}_direct_print_pdf",
                use_container_width=True,
            )
    except Exception as e:
        st.error(f"Print Report error: {e}")
        st.exception(e)




# -----------------------------
# Fill Date & Time from Excel + PDF
# -----------------------------
def _norm_paper_code(v):
    return re.sub(r"[^A-Za-z0-9]", "", str(v or "")).upper()


def _norm_subject_text(v):
    """Normalize subject text only for safe formatting differences."""
    s = str(v or "")
    s = s.replace("&", " and ")
    s = re.sub(r"\bA\.?\s*M\.?\b", "AM", s, flags=re.I)
    s = re.sub(r"\bP\.?\s*M\.?\b", "PM", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


def _subject_variants(v):
    """
    Produce deterministic subject-name variants.

    Allowed normalization:
      - case/space/punctuation differences
      - '&' == 'and'
      - remove presentation-only labels such as (CC), (DSE), (AEC)
      - remove one leading paper/subject identifier when it is clearly a
        separate prefix (e.g. PE1, CPS1, CA-111)

    No fuzzy matching is used.
    """
    s = str(v or "")
    variants = [s]

    # Presentation labels in the PDF are not part of the actual subject name.
    s2 = re.sub(
        r"\s*\(\s*(?:C\s*C|DSE|AEC|DSEC)\s*\)\s*\)?",
        " ",
        s,
        flags=re.I,
    )
    s2 = re.sub(r"\s+", " ", s2).strip()
    variants.append(s2)

    # Some Excel PaperName values begin with a paper identifier that is not
    # present in PDF Subject Name. Remove only that leading identifier.
    for v2 in list(variants):
        m = re.match(
            r"^\s*[A-Za-z]{1,8}[- ]?\d{1,5}(?:\s*[A-Za-z])?\s*[:.)-]?\s+(.+)$",
            v2,
        )
        if m:
            variants.append(m.group(1))

    return {
        _norm_subject_text(v)
        for v in variants
        if _norm_subject_text(v)
    }


def _parse_pdf_date(text):
    patterns = [
        r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b",
        r"\b(\d{1,2})[- ]([A-Za-z]{3,9})[- ](\d{2,4})\b",
        r"\b(\d{1,2})[./-](\d{1,2})\b",
    ]
    months = {
        "JAN": 1, "JANUARY": 1, "FEB": 2, "FEBRUARY": 2,
        "MAR": 3, "MARCH": 3, "APR": 4, "APRIL": 4,
        "MAY": 5, "JUN": 6, "JUNE": 6, "JUL": 7, "JULY": 7,
        "AUG": 8, "AUGUST": 8, "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
        "OCT": 10, "OCTOBER": 10, "NOV": 11, "NOVEMBER": 11,
        "DEC": 12, "DECEMBER": 12,
    }
    for pat in patterns:
        for m in re.finditer(pat, str(text), flags=re.I):
            try:
                if m.lastindex >= 3:
                    a, b, c = m.group(1), m.group(2), m.group(3)
                    if b.upper() in months:
                        d, mo, y = int(a), months[b.upper()], int(c)
                    else:
                        d, mo, y = int(a), int(b), int(c)
                    if y < 100:
                        y += 2000
                else:
                    d, mo = int(m.group(1)), int(m.group(2))
                    y = date.today().year
                if 1 <= d <= 31 and 1 <= mo <= 12:
                    return date(y, mo, d)
            except Exception:
                continue
    return None


def _format_pdf_clock(token):
    """Format one clock value as compact AM/PM text, e.g. 9 AM or 11:30 AM."""
    s = re.sub(r"\s+", " ", str(token).strip().upper())
    s = (
        s.replace("A.M.", "AM").replace("P.M.", "PM")
        .replace("A.M", "AM").replace("P.M", "PM")
        .replace("NOON", "PM")
    )
    m = re.fullmatch(r"(\d{1,2})(?:(?::|\.)(\d{1,2}))?\s*(AM|PM)", s)
    if not m:
        return s
    hour = int(m.group(1))
    minute = m.group(2)
    suffix = m.group(3)
    if minute is None or int(minute) == 0:
        return f"{hour} {suffix}"
    return f"{hour}:{int(minute):02d} {suffix}"


def _parse_pdf_time(text):
    """Normalize PDF time to display like '9 AM to 11 AM'."""
    upper = re.sub(r"\s+", " ", str(text).upper())
    upper = (
        upper.replace("A.M.", "AM").replace("P.M.", "PM")
        .replace("A.M", "AM").replace("P.M", "PM")
        .replace("NOON", "PM")
    )
    ms = list(re.finditer(
        r"\b\d{1,2}(?::|\.)?\d{0,2}\s*(?:AM|PM)\b",
        upper,
    ))
    if len(ms) >= 2:
        for a, b in zip(ms, ms[1:]):
            between = upper[a.end():b.start()]
            if re.search(r"\bTO\b|-", between):
                return f"{_format_pdf_clock(a.group(0))} to {_format_pdf_clock(b.group(0))}"
    if ms:
        return _format_pdf_clock(ms[0].group(0))
    return None


def _pdf_lines(page):
    words = page.get_text("words")
    groups = {}
    for w in words:
        y = round(float(w[1]) / 3.0) * 3.0
        groups.setdefault(y, []).append(w)
    lines = []
    for y, items in groups.items():
        items.sort(key=lambda x: x[0])
        lines.append((y, " ".join(str(x[4]) for x in items)))
    lines.sort(key=lambda x: x[0])
    return [txt for _, txt in lines if txt.strip()]


def _pdf_plain_lines(page):
    """Fallback line extraction preserving the PDF's logical text order."""
    return [
        re.sub(r"\s+", " ", line).strip()
        for line in page.get_text().splitlines()
        if line.strip()
    ]


def _find_nearby_pdf_date(lines, idx, radius=18):
    best = None
    for j in range(max(0, idx - radius), min(len(lines), idx + radius + 1)):
        d = _parse_pdf_date(lines[j])
        if d:
            candidate = (abs(j - idx), j, d)
            if best is None or candidate < best:
                best = candidate
    return best[2] if best else None


def _find_nearby_pdf_time(lines, idx, radius=18):
    best = None
    for j in range(max(0, idx - radius), min(len(lines), idx + radius + 1)):
        context = " ".join(lines[max(0, j - 1): min(len(lines), j + 2)])
        t = _parse_pdf_time(context)
        if t:
            candidate = (abs(j - idx), j, t)
            if best is None or candidate < best:
                best = candidate
    return best[2] if best else None


def _pdf_code_hits(lines, wanted):
    """Find requested Computer Codes as standalone numeric tokens."""
    hits = []
    code_re = re.compile(r"(?<!\d)(\d{4,12})(?!\d)")
    wanted = set(wanted or [])
    for i, line in enumerate(lines):
        for m in code_re.finditer(str(line)):
            code = _norm_paper_code(m.group(1))
            if code in wanted:
                hits.append((i, code, m.start(), m.end()))
    return hits


def extract_pdf_paper_map(pdf_bytes, excel_rows):
    """Match Excel PaperCode ONLY to PDF Computer Code.

    Once an exact code match is found, the Date and the COMPLETE Time range
    belonging to that PDF entry are returned. PaperName/SubjectName are not
    used for matching. Conflicting Date/Time pairs for the same code are
    treated as ambiguous rather than guessed.
    """
    if fitz is None:
        raise RuntimeError("PDF library (PyMuPDF) is not installed.")

    wanted_codes = {_norm_paper_code(code) for _, code, _ in excel_rows if _norm_paper_code(code)}
    code_candidates = {code: set() for code in wanted_codes}

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page in doc:
            lines = _pdf_plain_lines(page)
            for idx, code, _, _ in _pdf_code_hits(lines, wanted_codes):
                    # Only accept the exact code match. Date and the complete
                    # time range are taken from the same nearby PDF entry.
                    # _find_nearby_pdf_time preserves ranges such as
                    # "9 AM to 11 AM" instead of reducing them to "9 AM".
                    date_value = _find_nearby_pdf_date(lines, idx, radius=18)
                    time_value = _find_nearby_pdf_time(lines, idx, radius=18)
                    if date_value and time_value:
                        code_candidates[code].add((date_value, time_value))
    finally:
        doc.close()

    # Map by exact normalized code; no subject/name condition.
    code_pairs = {code: pairs for code, pairs in code_candidates.items() if pairs}
    row_map = {}
    raw_matches = {}
    for row, code, _ in excel_rows:
        ncode = _norm_paper_code(code)
        pairs = code_pairs.get(ncode, set())
        raw_matches[row] = [{"code": ncode, "date": d, "time": t} for d, t in sorted(pairs)]
        row_map[row] = next(iter(pairs)) if len(pairs) == 1 else None
    return row_map, raw_matches

def fill_excel_date_time_from_pdf(excel_bytes, pdf_bytes):
    wb = load_workbook(BytesIO(excel_bytes), data_only=False)
    ws = get_main_sheet(wb)
    h = find_header_row(ws, {"date", "time"})
    if not h:
        raise ValueError("Excel मध्ये Date आणि Time headings सापडल्या नाहीत.")

    date_col = find_col(ws, {"date"}, h)
    time_col = find_col(ws, {"time"}, h)
    code_col = find_col(ws, {"paper code", "papercode", "paper code."}, h)
    if not all([date_col, time_col, code_col]):
        raise ValueError("Excel मध्ये Date, Time किंवा Paper Code column सापडला नाही.")

    excel_rows = []
    for r in range(h + 1, ws.max_row + 1):
        code = ws.cell(r, code_col).value
        if code not in (None, ""):
            excel_rows.append((r, code, None))

    row_map, raw_matches = extract_pdf_paper_map(pdf_bytes, excel_rows)
    ok = errors = 0
    red_font = Font(color="FF0000", bold=True)

    for r, _, _ in excel_rows:
        dc, tc = ws.cell(r, date_col), ws.cell(r, time_col)
        pair = row_map.get(r)
        if pair:
            dt_value, tm_value = pair
            dc.value = dt_value
            dc.number_format = "dd/mm/yyyy"
            tc.value = tm_value
            dc.font = copy(dc.font)
            tc.font = copy(tc.font)
            ok += 1
        else:
            dc.value = "ERROR"
            tc.value = "ERROR"
            dc.font = copy(red_font)
            tc.font = copy(red_font)
            errors += 1

    merge_contiguous_date_time(ws, date_col, time_col, h)
    return wb, ok, errors, raw_matches


# -----------------------------
# Final KD-EXAM application flow
# -----------------------------
# IMPORTANT: every tab starts from the ORIGINAL uploaded Excel stored in
# session_state.original_bytes. Derived workbooks are created only inside the tab.

def original_source():
    return st.session_state.get("original_bytes")


def validate_original_upload(source):
    wb = load_workbook(BytesIO(source), data_only=False)
    ws = get_main_sheet(wb)
    h = find_header_row(ws, {"date", "time"})
    if not h:
        raise ValueError("Date आणि Time headings सापडले नाहीत.")
    missing = []
    for name in ("Date", "Time", "Paper Code"):
        if not find_col(ws, {name.lower()}, h):
            missing.append(name)
    if missing:
        raise ValueError("Required headings missing: " + ", ".join(missing))
    return wb


@st.cache_data(show_spinner=False, ttl=600)
def _cached_dform_bytes(source, start_date, end_date, selected_classes_tuple):
    wb = load_workbook(BytesIO(source), data_only=False)
    filter_workbook_by_date_and_class(wb, start_date, end_date, list(selected_classes_tuple))
    ws = get_main_sheet(wb)
    h = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, h)
    time_col = find_col(ws, {"time"}, h)
    merge_contiguous_date_time(ws, date_col, time_col, h)
    wb = process_formatting(wb)
    wb = process_dform(wb)
    apply_excel_print_formatting(wb)
    return wb_bytes(wb)


@st.cache_data(show_spinner=False, ttl=600)
def _cached_formatted_bytes(source, start_date, end_date, selected_classes_tuple, do_dform=False):
    wb = load_workbook(BytesIO(source), data_only=False)
    filter_workbook_by_date_and_class(wb, start_date, end_date, list(selected_classes_tuple))
    ws = get_main_sheet(wb)
    h = find_header_row(ws, {"date", "time"}) or 1
    dc = find_col(ws, {"date"}, h)
    tc = find_col(ws, {"time"}, h)
    if dc and tc:
        merge_contiguous_date_time(ws, dc, tc, h)
    wb = process_formatting(wb)
    if do_dform:
        wb = process_dform(wb)
    apply_excel_print_formatting(wb)
    return wb_bytes(wb)


def fresh_dform(source, start_date, end_date, selected_classes):
    """Build D Form fresh from ORIGINAL upload for the current selection."""
    data = _cached_dform_bytes(
        source, start_date, end_date, tuple(selected_classes or [])
    )
    return load_workbook(BytesIO(data), data_only=False)


def fresh_formatted(source, start_date=None, end_date=None, selected_classes=None, do_dform=False):
    if start_date is None or end_date is None:
        return load_workbook(BytesIO(source), data_only=False)
    data = _cached_formatted_bytes(
        source, start_date, end_date, tuple(selected_classes or []), do_dform
    )
    return load_workbook(BytesIO(data), data_only=False)


def style_red_checkbox():
    st.markdown(
        """
        <style>
        input[type='checkbox'] { accent-color: #d00000 !important; }
        div[data-testid='stCheckbox'] label p { color: #000000 !important; font-weight: 500; }
        </style>
        """,
        unsafe_allow_html=True,
    )


CALCULATED_REPORT_HEADERS = {
    "total students",
    "no. of blocks",
    "supervisors",
    "int. squad",
    "int. sr. supervisor",
    "ext. sr. supervisor",
    "sessions",
    "entrance code",
}

def is_calculated_report_header(value):
    s = re.sub(r"\s+", " ", str(value or "").strip().lower())
    return s in CALCULATED_REPORT_HEADERS

def copy_selected_columns_with_merges(src_ws, dst_ws, header_row, selected_cols):
    """Copy selected source columns while preserving vertical merge structure."""
    for dst_col, src_col in enumerate(selected_cols, start=1):
        copy_col_dimension(src_ws, src_col, dst_col)
        for r in range(1, src_ws.max_row + 1):
            src_cell = src_ws.cell(r, src_col)
            dst_cell = dst_ws.cell(r, dst_col)
            if isinstance(dst_cell, MergedCell):
                continue
            copy_cell_style(src_cell, dst_cell)
            dst_cell.value = None if isinstance(src_cell, MergedCell) else src_cell.value

        for rng in list(src_ws.merged_cells.ranges):
            if rng.min_col == src_col and rng.max_col == src_col:
                dst_ws.merge_cells(
                    start_row=rng.min_row, start_column=dst_col,
                    end_row=rng.max_row, end_column=dst_col
                )

    wrap_all_headers(dst_ws, header_row)

def vertical_heading_selector(headers, prefix):
    """Vertical Column checkbox list with Select All."""
    return vertical_checkbox_selector(
        headers,
        prefix,
        label="Column",
        collapsed=False,
    )


def selected_date_from_source(source, prefix):
    dates = available_report_dates(source)
    if not dates:
        return None
    return st.selectbox(
        "Date",
        dates,
        format_func=lambda d: d.strftime("%d/%m/%Y"),
        key=f"{prefix}_date",
    )


def add_duration_tables(summary_ws, by_date_times, start_row):
    """Append the requested 2-hour and 3-hour tables below existing Session Summary."""
    def parse_clock(text, inherited_ap=None):
        s = str(text).strip().lower().replace(".", ":")
        s = re.sub(r"\s+", " ", s)
        m = re.search(r"(\d{1,2})(?::(\d{1,2}))?\s*(am|pm)?", s)
        if not m:
            return None
        hh = int(m.group(1))
        mm = int(m.group(2) or 0)
        ap = m.group(3) or inherited_ap
        if mm > 59 or hh > 23:
            return None
        if ap == "pm" and hh != 12:
            hh += 12
        elif ap == "am" and hh == 12:
            hh = 0
        return hh * 60 + mm

    def duration_minutes(t):
        s = str(t).strip().lower().replace("–", "-").replace("—", "-")
        s = re.sub(r"\s+to\s+", "-", s)
        s = re.sub(r"\s*-\s*", "-", s)
        parts = s.split("-", 1)
        if len(parts) != 2:
            return None

        left, right = parts[0].strip(), parts[1].strip()
        right_ap_match = re.search(r"\b(am|pm)\b", right)
        left_ap_match = re.search(r"\b(am|pm)\b", left)
        right_ap = right_ap_match.group(1) if right_ap_match else None
        left_ap = left_ap_match.group(1) if left_ap_match else right_ap

        a = parse_clock(left, left_ap)
        b = parse_clock(right, right_ap)
        if a is None or b is None:
            return None
        if b < a:
            b += 24 * 60
        return b - a

    two = []
    three = []
    for d, times in sorted(by_date_times.items()):
        c2 = sum(1 for t in times if (lambda x: x is not None and 1 <= x <= 120)(duration_minutes(t)))
        c3 = sum(1 for t in times if (lambda x: x is not None and x > 120)(duration_minutes(t)))
        if c2:
            two.append((d, c2))
        if c3:
            three.append((d, c3))

    row = max(start_row + 2, summary_ws.max_row + 2)
    summary_ws.cell(row, 1).value = "2 Hours"
    summary_ws.cell(row, 1).font = Font(bold=True, size=12)
    row += 1
    summary_ws.cell(row, 1).value = "Dates (dd/mm)"
    summary_ws.cell(row, 2).value = "Count"
    for c in (1, 2): summary_ws.cell(row, c).font = Font(bold=True)
    row += 1
    if two:
        for d, count in two:
            summary_ws.cell(row, 1).value = d.strftime("%d/%m")
            summary_ws.cell(row, 2).value = count
            row += 1
    else:
        summary_ws.cell(row, 1).value = "No sessions"
        row += 1

    row += 1
    summary_ws.cell(row, 1).value = "3 Hours"
    summary_ws.cell(row, 1).font = Font(bold=True, size=12)
    row += 1
    summary_ws.cell(row, 1).value = "Dates (dd/mm)"
    summary_ws.cell(row, 2).value = "Count"
    for c in (1, 2): summary_ws.cell(row, c).font = Font(bold=True)
    row += 1
    if three:
        for d, count in three:
            summary_ws.cell(row, 1).value = d.strftime("%d/%m")
            summary_ws.cell(row, 2).value = count
            row += 1
    else:
        summary_ws.cell(row, 1).value = "No sessions"

    for r in range(start_row, summary_ws.max_row + 1):
        for c in range(1, 3):
            summary_ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    summary_ws.column_dimensions["A"].width = max(summary_ws.column_dimensions["A"].width or 12, 20)
    summary_ws.column_dimensions["B"].width = max(summary_ws.column_dimensions["B"].width or 12, 14)


def session_times_by_date(wb):
    ws = get_main_sheet(wb)
    h = find_header_row(ws, {"date", "time"}) or 1
    dc = find_col(ws, {"date"}, h); tc = find_col(ws, {"time"}, h)
    out = {}
    cur = None
    for r in range(h + 1, ws.max_row + 1):
        dv = ws.cell(r, dc).value
        if dv not in (None, ""):
            try: cur = datetime.strptime(date_key(dv), "%Y-%m-%d").date()
            except Exception: cur = None
        tv = ws.cell(r, tc).value
        if cur and tv not in (None, ""):
            out.setdefault(cur, []).append(tv)
    return out


# -------- Tabs / File Inputs --------
t0, t1, t2, t3, t4 = st.tabs([
    "1. Fill Date & Time", "2. Schedule for Squad", "3. Session Count", "4. D Form", "5. Dashboard"
])

with t0:
    st.subheader("Fill Date & Time")
    st.caption("Excel PaperCode ↔ PDF Computer Code → त्याच PDF entry मधील Date + पूर्ण Time Range")
    excel_input = st.file_uploader("1. Input Excel", type=["xlsx", "xlsm"], key="fill_excel_input")
    pdf_input = st.file_uploader("2. Input PDF", type=["pdf"], key="fill_pdf_input")
    if excel_input is not None and pdf_input is not None:
        try:
            excel_data = excel_input.getvalue(); pdf_data = pdf_input.getvalue()
            fill_sig = hashlib.sha256(excel_data + pdf_data).hexdigest()
            if st.session_state.get("fill_datetime_signature") != fill_sig:
                with st.spinner("PaperCode आणि PDF Computer Code exact match करून Date + पूर्ण Time Range घेत आहे..."):
                    filled_wb, ok_count, err_count, _ = fill_excel_date_time_from_pdf(excel_data, pdf_data)
                    filled_bytes = wb_bytes(filled_wb)
                st.session_state.filled_datetime_bytes = filled_bytes
                st.session_state.fill_datetime_signature = fill_sig
                st.session_state.fill_ok_count = ok_count
                st.session_state.fill_error_count = err_count
            else:
                filled_bytes = st.session_state.get("filled_datetime_bytes")
                ok_count = st.session_state.get("fill_ok_count", 0)
                err_count = st.session_state.get("fill_error_count", 0)
            st.success(f"Date/Time भरले: {ok_count} rows | ERROR: {err_count} rows")
            if filled_bytes:
                st.download_button("⬇️ Download Filled Excel", data=filled_bytes,
                    file_name="KD_EXAM_Filled_Date_Time.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="filled_datetime_download", use_container_width=True)
        except Exception as e:
            st.error(f"Fill Date & Time error: {e}")
            st.exception(e)
    else:
        st.info("कृपया Excel आणि PDF दोन्ही upload करा.")

# -------- 1. SCHEDULE FOR SQUAD --------
with t1:
    st.subheader("Schedule for Squad")
    st.caption("या tab मध्ये Excel upload करा. त्यावर Date/Time merge, formatting आणि No. of Blocks पर्यंत calculation होईल.")
    schedule_input = st.file_uploader(
        "Schedule for Squad साठी Excel फाईल upload करा",
        type=["xlsx", "xlsm"], key="schedule_excel_input"
    )
    if schedule_input is None:
        st.info("कृपया Schedule for Squad साठी Excel फाईल upload करा. हीच processed फाईल पुढील Session Count आणि D Form tabs मध्ये input म्हणून वापरली जाईल.")
    else:
        try:
            schedule_raw = schedule_input.getvalue()
            schedule_sig = hashlib.sha256(schedule_raw).hexdigest()
            if st.session_state.get("schedule_upload_signature") != schedule_sig:
                with st.spinner("Schedule तयार करत आहे: Date/Time merge + formatting + calculation..."):
                    swb = load_workbook(BytesIO(schedule_raw), data_only=False)
                    sws = get_main_sheet(swb)
                    sh = find_header_row(sws, {"date", "time"}) or 1
                    dc = find_col(sws, {"date"}, sh); tc = find_col(sws, {"time"}, sh)
                    if not dc or not tc:
                        raise ValueError("Uploaded Excel मध्ये Date आणि Time columns आवश्यक आहेत.")
                    start_all, end_all = all_date_range(schedule_raw)
                    if start_all is None or end_all is None:
                        raise ValueError("Uploaded Excel मध्ये valid Date data सापडला नाही.")
                    filter_workbook_by_date_and_class(swb, start_all, end_all, [])
                    sws = get_main_sheet(swb)
                    sh = find_header_row(sws, {"date", "time"}) or 1
                    dc = find_col(sws, {"date"}, sh); tc = find_col(sws, {"time"}, sh)
                    merge_contiguous_date_time(sws, dc, tc, sh)
                    swb = process_formatting(swb)
                    swb = process_dform(swb)
                    apply_excel_print_formatting(swb)
                    processed = wb_bytes(swb)
                st.session_state.schedule_upload_signature = schedule_sig
                st.session_state.schedule_processed_bytes = processed
            else:
                processed = st.session_state.get("schedule_processed_bytes")

            # Full processed workbook is the common input for the next two tabs.
            pwb = load_workbook(BytesIO(processed), data_only=False)
            pws = get_main_sheet(pwb)
            ph = find_header_row(pws, {"date", "time"}) or 1
            pblock = find_col(pws, {"no. of blocks"}, ph)
            if not pblock:
                raise ValueError("No. of Blocks column सापडला नाही.")

            st.success("Schedule for Squad processing पूर्ण झाले. Date/Time merge, formatting आणि No. of Blocks पर्यंत calculation झाले आहे.")
            all_preview = [[pws.cell(r,c).value for c in range(1,pblock+1)]
                           for r in range(ph+1, min(pws.max_row, ph+40)+1)]
            st.dataframe(all_preview, use_container_width=True, hide_index=True)

            schedule_headers = [pws.cell(ph,c).value or f"Column {c}" for c in range(1,pblock+1)]
            selected_headers = vertical_heading_selector(schedule_headers, "squad_heading_v13")
            if selected_headers:
                wanted = {str(v) for v in selected_headers}
                selected_idx = [c for c in range(1,pblock+1)
                                if str(pws.cell(ph,c).value or f"Column {c}") in wanted]
                out = load_workbook(BytesIO(processed), data_only=False)
                src = get_main_sheet(out)
                new = out.create_sheet("Schedule for Squad")
                copy_selected_columns_with_merges(src, new, ph, selected_idx)
                for ws in list(out.worksheets):
                    if ws is not new: out.remove(ws)
                out.active = 0
                apply_excel_print_formatting(out)
                squad_bytes = wb_bytes(out)
                load_workbook(BytesIO(squad_bytes), data_only=False)
                st.download_button("Download Schedule for Squad", squad_bytes,
                    "KD_EXAM_Schedule_for_Squad.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="squad_download_v13", on_click="ignore")
            else:
                st.warning("किमान एक Column निवडा.")
        except Exception as e:
            st.error(f"Schedule for Squad error: {e}")
            st.exception(e)

# -------- 2. SESSION COUNT --------
with t2:
    st.subheader("Session Count")
    schedule_source = st.session_state.get("schedule_processed_bytes")
    if not schedule_source:
        st.info("पहिल्या Schedule for Squad tab मध्ये Excel upload करा. त्याच Excel वर Session Count तयार होईल.")
    else:
        st.caption("Input = Schedule for Squad मध्ये upload केलेली Excel फाईल (processed version)")
        try:
            start_date, end_date, selected_classes = date_range_controls(schedule_source, "session_count", include_class=True)
            if start_date is not None and end_date is not None:
                wb = fresh_dform(schedule_source, start_date, end_date, selected_classes)
                wb, total_dates, total_sessions = process_sessions(wb)
                summary_ws = wb["Session Summary"]
                by_date_times = session_times_by_date(wb)
                add_duration_tables(summary_ws, by_date_times, summary_ws.max_row)
                apply_excel_print_formatting(wb)
                session_bytes = wb_bytes(wb)
                load_workbook(BytesIO(session_bytes), data_only=False)
                st.success(f"Session Count: {total_dates} dates / {total_sessions} sessions")
                st.download_button("Download Session Count", session_bytes,
                    "KD_EXAM_Session_Count.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="session_count_download_v13", on_click="ignore")
        except Exception as e:
            st.error(f"Session Count error: {e}")
            st.exception(e)

# -------- 3. D FORM --------
with t3:
    st.subheader("D Form")
    schedule_source = st.session_state.get("schedule_processed_bytes")
    if not schedule_source:
        st.info("पहिल्या Schedule for Squad tab मध्ये Excel upload करा. त्याच Excel वर D Form तयार होईल.")
    else:
        st.caption("Input = Schedule for Squad मध्ये upload केलेली Excel फाईल (processed version)")
        try:
            dates = available_report_dates(schedule_source)
            if not dates:
                st.warning("Excel मध्ये Date data सापडला नाही.")
            else:
                selected_classes = vertical_class_selector(schedule_source, "dform", collapsed=True)
                start_date, end_date = all_date_range(schedule_source)
                if not selected_classes and available_report_classes(schedule_source):
                    st.warning("किमान एक Class निवडा.")
                else:
                    dform_wb = fresh_dform(schedule_source, start_date, end_date, selected_classes)
                    dform_bytes = wb_bytes(dform_wb)
                    st.success(f"D Form तयार — {start_date.strftime('%d/%m/%Y')} ते {end_date.strftime('%d/%m/%Y')}")
                    ws = get_main_sheet(dform_wb)
                    h = find_header_row(ws, {"date", "time"}) or 1
                    st.dataframe([[ws.cell(r,c).value for c in range(1,ws.max_column+1)]
                                  for r in range(h+1,min(ws.max_row,h+21)+1)],
                                 use_container_width=True, hide_index=True)
                    st.download_button("Download D Form", dform_bytes, "KD_EXAM_D_Form.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dform_download_v13")
        except Exception as e:
            st.error(f"D Form error: {e}")
            st.exception(e)

# -------- 4. DASHBOARD --------
with t4:
    st.subheader("Dashboard")
    schedule_source = st.session_state.get("schedule_processed_bytes")
    if not schedule_source:
        st.info("पहिल्या Schedule for Squad tab मध्ये Excel upload करा.")
    else:
        st.caption("Schedule for Squad processed Excel वर आधारित Dashboard")
        try:
            current_date = date.today()
            st.markdown(f"### Date: {current_date.strftime('%d/%m/%Y')}")
            day_wb = fresh_dform(schedule_source, current_date, current_date, [])
            ws = get_main_sheet(day_wb); h = find_header_row(ws,{"date","time"}) or 1
            tc=find_col(ws,{"time"},h); bc=find_col(ws,{"no. of blocks"},h); sc=find_col(ws,{"supervisors"},h); dc=find_col(ws,{"date"},h)
            rows=[]; cur_date=None
            for r in range(h+1,ws.max_row+1):
                dv=ws.cell(r,dc).value
                if dv not in (None,""):
                    try: cur_date=datetime.strptime(date_key(dv),"%Y-%m-%d").date()
                    except Exception: cur_date=None
                if cur_date != current_date: continue
                tv=ws.cell(r,tc).value if tc else None; bv=ws.cell(r,bc).value if bc else None; sv=ws.cell(r,sc).value if sc else None
                if tv not in (None,""): rows.append((str(tv),bv,sv))
            display=[]; seen=set()
            for row in rows:
                if row not in seen: seen.add(row); display.append(row)
            if display: st.dataframe(display,use_container_width=True,hide_index=True,column_config={0:"Time",1:"No. of Blocks",2:"Supervisors"})
            else: st.info("आजच्या तारखेची माहिती uploaded Schedule Excel मध्ये उपलब्ध नाही.")
            st.download_button("Download Dashboard", wb_bytes(day_wb), "KD_EXAM_Dashboard.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dashboard_download_v13")
        except Exception as e:
            st.error(f"Dashboard error: {e}")
            st.exception(e)


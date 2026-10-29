from __future__ import annotations

import sys
import os
import io
import urllib.request
from datetime import date

import streamlit as st
from supabase import create_client, Client
from PIL import Image, ImageOps
import openpyxl

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
)

from openpyxl.drawing.image import Image as OpenPyXLImage


# ==============================================================================
# PATHS AND IMPORTS
# ==============================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import ai_service


# ==============================================================================
# CONSTANTS
# ==============================================================================

QC_REPORTS_BUCKET = "qc-reports"

EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)


# ==============================================================================
# EXCEL HELPER FUNCTIONS
# ==============================================================================

def apply_range_style(
    ws,
    min_col,
    min_row,
    max_col,
    max_row,
    border=None,
    fill=None,
):
    """Apply borders and fills to all cells in a range."""

    for row in ws.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
    ):
        for cell in row:
            if border:
                cell.border = border

            if fill:
                cell.fill = fill


# ==============================================================================
# EXCEL REPORT GENERATOR
# ==============================================================================

def create_full_qc_report(
    report_data: dict,
    photos_by_section: dict,
    spec_data: dict = None,
    measured_samples: list = None,
) -> bytes:

    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "QC Inspection Report"
    ws.views.sheetView[0].showGridLines = True

    # --------------------------------------------------------------------------
    # COLORS
    # --------------------------------------------------------------------------

    c_brand = "1E293B"
    c_header = "334155"
    c_sub = "475569"
    c_border = "CBD5E1"
    c_meta_bg = "F8FAFC"

    c_fail_bg = "FEE2E2"
    c_fail_fg = "991B1B"

    c_pass_bg = "DCFCE7"
    c_pass_fg = "166534"

    c_pend_bg = "FEF3C7"
    c_pend_fg = "92400E"

    thin_border = Border(
        left=Side(style="thin", color=c_border),
        right=Side(style="thin", color=c_border),
        top=Side(style="thin", color=c_border),
        bottom=Side(style="thin", color=c_border),
    )

    # --------------------------------------------------------------------------
    # COLUMN WIDTHS
    # --------------------------------------------------------------------------

    ws.column_dimensions["A"].width = 3

    for col_letter in [
        "B", "C", "D", "E",
        "F", "G", "H", "I",
    ]:
        ws.column_dimensions[col_letter].width = 15

    for col_letter in [
        "J", "K", "L", "M", "N", "O",
    ]:
        ws.column_dimensions[col_letter].width = 14

    # --------------------------------------------------------------------------
    # 1. HEADER
    # --------------------------------------------------------------------------

    ws.merge_cells("B2:I2")
    ws.row_dimensions[2].height = 40

    company_title = report_data.get(
        "company_name",
        "BIZNESKIT",
    ).upper()

    h_cell = ws["B2"]

    h_cell.value = (
        f"{company_title} — "
        "QUALITY CONTROL INSPECTION REPORT"
    )

    h_cell.fill = PatternFill(
        start_color=c_brand,
        end_color=c_brand,
        fill_type="solid",
    )

    h_cell.font = Font(
        name="Calibri",
        size=13,
        bold=True,
        color="FFFFFF",
    )

    h_cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    apply_range_style(
        ws,
        2, 2, 9, 2,
        border=thin_border,
    )

    # --------------------------------------------------------------------------
    # 2. INSPECTION INFORMATION
    # --------------------------------------------------------------------------

    res_status = str(
        report_data.get("result_status", "PENDING")
    ).upper()

    meta_rows = [
        (
            "REPORT NO:",
            report_data.get("report_no", "QC-001"),
            "INSPECTION DATE:",
            str(
                report_data.get(
                    "inspection_date",
                    date.today().isoformat(),
                )
            ),
        ),
        (
            "CLIENT:",
            report_data.get("client_name", "-"),
            "INSPECTOR ID:",
            str(report_data.get("inspector_id", "-")),
        ),
        (
            "SUPPLIER / FACTORY:",
            report_data.get("supplier_name", "-"),
            "PO / ORDER NO:",
            str(report_data.get("order_no", "-")),
        ),
        (
            "ITEM / SKU:",
            (
                f"{report_data.get('item_name', '-')} / "
                f"{report_data.get('sku', '-')}"
            ),
            "OVERALL RESULT:",
            res_status,
        ),
    ]

    curr_row = 4

    for r_data in meta_rows:

        ws.row_dimensions[curr_row].height = 22

        # First label
        c1 = ws.cell(
            row=curr_row,
            column=2,
            value=r_data[0],
        )

        c1.fill = PatternFill(
            start_color=c_meta_bg,
            fill_type="solid",
        )

        c1.font = Font(
            name="Calibri",
            size=9,
            bold=True,
            color=c_sub,
        )

        c1.border = thin_border

        # First value
        ws.merge_cells(
            start_row=curr_row,
            start_column=3,
            end_row=curr_row,
            end_column=5,
        )

        c2 = ws.cell(
            row=curr_row,
            column=3,
            value=str(r_data[1]),
        )

        c2.font = Font(
            name="Calibri",
            size=9,
            bold=True,
            color="0F172A",
        )

        c2.alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=True,
        )

        apply_range_style(
            ws,
            3, curr_row, 5, curr_row,
            border=thin_border,
        )

        # Second label
        c3 = ws.cell(
            row=curr_row,
            column=6,
            value=r_data[2],
        )

        c3.fill = PatternFill(
            start_color=c_meta_bg,
            fill_type="solid",
        )

        c3.font = Font(
            name="Calibri",
            size=9,
            bold=True,
            color=c_sub,
        )

        c3.border = thin_border

        # Second value
        ws.merge_cells(
            start_row=curr_row,
            start_column=7,
            end_row=curr_row,
            end_column=9,
        )

        c4 = ws.cell(
            row=curr_row,
            column=7,
            value=str(r_data[3]),
        )

        apply_range_style(
            ws,
            7, curr_row, 9, curr_row,
            border=thin_border,
        )

        if r_data[2] == "OVERALL RESULT:":

            c4.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

            if res_status == "PASSED":

                c4.fill = PatternFill(
                    start_color=c_pass_bg,
                    fill_type="solid",
                )

                c4.font = Font(
                    name="Calibri",
                    size=11,
                    bold=True,
                    color=c_pass_fg,
                )

            elif res_status == "REJECTED":

                c4.fill = PatternFill(
                    start_color=c_fail_bg,
                    fill_type="solid",
                )

                c4.font = Font(
                    name="Calibri",
                    size=11,
                    bold=True,
                    color=c_fail_fg,
                )

            else:

                c4.fill = PatternFill(
                    start_color=c_pend_bg,
                    fill_type="solid",
                )

                c4.font = Font(
                    name="Calibri",
                    size=11,
                    bold=True,
                    color=c_pend_fg,
                )

        else:

            c4.font = Font(
                name="Calibri",
                size=9,
                color="0F172A",
            )

            c4.alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=True,
            )

        curr_row += 1

    curr_row += 1

    # --------------------------------------------------------------------------
    # 3. GARMENT MEASUREMENT SPECIFICATION
    # --------------------------------------------------------------------------

    if spec_data and measured_samples:

        ws.merge_cells(
            start_row=curr_row,
            start_column=2,
            end_row=curr_row,
            end_column=9,
        )

        ws.row_dimensions[curr_row].height = 24

        sec_cell = ws.cell(
            row=curr_row,
            column=2,
            value="GARMENT MEASUREMENT & TOLERANCE SPEC SHEET",
        )

        sec_cell.fill = PatternFill(
            start_color=c_header,
            fill_type="solid",
        )

        sec_cell.font = Font(
            name="Calibri",
            size=10,
            bold=True,
            color="FFFFFF",
        )

        sec_cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            indent=1,
        )

        apply_range_style(
            ws,
            2, curr_row, 9, curr_row,
            border=thin_border,
        )

        curr_row += 1

        ws.row_dimensions[curr_row].height = 30

        spec_headers = [
            (2, 2, "Code"),
            (3, 5, "POM Description"),
            (6, 6, "Tol (-)"),
            (7, 7, "Tol (+)"),
        ]

        for s_col, e_col, h_name in spec_headers:

            if s_col != e_col:
                ws.merge_cells(
                    start_row=curr_row,
                    start_column=s_col,
                    end_row=curr_row,
                    end_column=e_col,
                )

            c = ws.cell(
                row=curr_row,
                column=s_col,
                value=h_name,
            )

            c.fill = PatternFill(
                start_color=c_sub,
                fill_type="solid",
            )

            c.font = Font(
                name="Calibri",
                size=9,
                bold=True,
                color="FFFFFF",
            )

            c.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

            apply_range_style(
                ws,
                s_col, curr_row, e_col, curr_row,
                border=thin_border,
            )

        sample_col = 8

        for sample in measured_samples:

            ws.cell(
                row=curr_row,
                column=sample_col,
                value=(
                    f"#{sample.get('sample_no', '')} "
                    f"({sample.get('size', '')}) Act"
                ),
            )

            ws.cell(
                row=curr_row,
                column=sample_col,
            ).font = Font(
                name="Calibri",
                size=8,
                bold=True,
                color="FFFFFF",
            )

            ws.cell(
                row=curr_row,
                column=sample_col,
            ).fill = PatternFill(
                start_color=c_sub,
                fill_type="solid",
            )

            ws.cell(
                row=curr_row,
                column=sample_col,
            ).alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

            ws.cell(
                row=curr_row,
                column=sample_col,
            ).border = thin_border

            ws.cell(
                row=curr_row,
                column=sample_col + 1,
                value="Diff",
            )

            ws.cell(
                row=curr_row,
                column=sample_col + 1,
            ).font = Font(
                name="Calibri",
                size=8,
                bold=True,
                color="FFFFFF",
            )

            ws.cell(
                row=curr_row,
                column=sample_col + 1,
            ).fill = PatternFill(
                start_color=c_sub,
                fill_type="solid",
            )

            ws.cell(
                row=curr_row,
                column=sample_col + 1,
            ).alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

            ws.cell(
                row=curr_row,
                column=sample_col + 1,
            ).border = thin_border

            sample_col += 2

        curr_row += 1

        for pom in spec_data.get("poms", []):

            ws.row_dimensions[curr_row].height = 22

            code = pom.get("code", "")

            code_cell = ws.cell(
                row=curr_row,
                column=2,
                value=code,
            )

            code_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

            code_cell.border = thin_border

            ws.merge_cells(
                start_row=curr_row,
                start_column=3,
                end_row=curr_row,
                end_column=5,
            )

            desc_cell = ws.cell(
                row=curr_row,
                column=3,
                value=pom.get("description", ""),
            )

            desc_cell.alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=True,
            )

            apply_range_style(
                ws,
                3, curr_row, 5, curr_row,
                border=thin_border,
            )

            tol_minus = pom.get("tol_minus", 1.0)
            tol_plus = pom.get("tol_plus", 1.0)

            c_minus = ws.cell(
                row=curr_row,
                column=6,
                value=f"-{tol_minus}",
            )

            c_plus = ws.cell(
                row=curr_row,
                column=7,
                value=f"+{tol_plus}",
            )

            for cell in [c_minus, c_plus]:

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                )

                cell.border = thin_border

            s_col = 8

            for sample in measured_samples:

                item = next(
                    (
                        x
                        for x in sample.get("items", [])
                        if x.get("code") == code
                    ),
                    None,
                )

                if item:

                    c_act = ws.cell(
                        row=curr_row,
                        column=s_col,
                        value=item.get("actual", "-"),
                    )

                    c_diff = ws.cell(
                        row=curr_row,
                        column=s_col + 1,
                        value=item.get("diff", 0),
                    )

                    c_act.border = thin_border
                    c_diff.border = thin_border

                    c_act.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                    )

                    c_diff.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                    )

                    c_act.font = Font(
                        name="Calibri",
                        size=9,
                    )

                    if item.get("status") == "FAIL":

                        c_diff.fill = PatternFill(
                            start_color=c_fail_bg,
                            fill_type="solid",
                        )

                        c_diff.font = Font(
                            name="Calibri",
                            size=9,
                            bold=True,
                            color=c_fail_fg,
                        )

                    else:

                        c_diff.fill = PatternFill(
                            start_color=c_pass_bg,
                            fill_type="solid",
                        )

                        c_diff.font = Font(
                            name="Calibri",
                            size=9,
                            bold=True,
                            color=c_pass_fg,
                        )

                s_col += 2

            curr_row += 1

        curr_row += 1

    # --------------------------------------------------------------------------
    # 4. PHOTO SECTIONS
    # --------------------------------------------------------------------------

    section_titles = {
        "products": "PRODUCT WORKMANSHIP PHOTOS",
        "packing": "PACKAGING & CARTON DETAILS",
        "labels": "LABELS & BARCODES VERIFICATION",
        "defects": "DEFECT DETAILS & LOG",
    }

    slots = [
        (2, 5, "B"),
        (6, 9, "F"),
    ]

    for s_key, s_label in section_titles.items():

        photos = photos_by_section.get(s_key, [])

        if not photos:
            continue

        ws.merge_cells(
            start_row=curr_row,
            start_column=2,
            end_row=curr_row,
            end_column=9,
        )

        ws.row_dimensions[curr_row].height = 22

        t_cell = ws.cell(
            row=curr_row,
            column=2,
            value=s_label,
        )

        t_cell.fill = PatternFill(
            start_color=c_header,
            fill_type="solid",
        )

        t_cell.font = Font(
            name="Calibri",
            size=10,
            bold=True,
            color="FFFFFF",
        )

        t_cell.alignment = Alignment(
            horizontal="left",
            vertical="center",
            indent=1,
        )

        apply_range_style(
            ws,
            2, curr_row, 9, curr_row,
            border=thin_border,
        )

        curr_row += 1

        for i in range(0, len(photos), 2):

            photo_row = curr_row
            ws.row_dimensions[photo_row].height = 195

            for slot_idx, (sc, ec, col_letter) in enumerate(slots):

                p_idx = i + slot_idx

                if p_idx >= len(photos):
                    continue

                p_info = photos[p_idx]

                ws.merge_cells(
                    start_row=photo_row,
                    start_column=sc,
                    end_row=photo_row,
                    end_column=ec,
                )

                apply_range_style(
                    ws,
                    sc, photo_row, ec, photo_row,
                    border=thin_border,
                )

                img = Image.open(
                    io.BytesIO(p_info["bytes"])
                )

                img = ImageOps.exif_transpose(img)

                if img.mode in ("RGBA", "LA", "P"):

                    if img.mode == "P":
                        img = img.convert("RGBA")

                    if img.mode in ("RGBA", "LA"):

                        background = Image.new(
                            "RGB",
                            img.size,
                            (255, 255, 255),
                        )

                        background.paste(
                            img,
                            mask=img.getchannel("A"),
                        )

                        img = background

                    else:
                        img = img.convert("RGB")

                elif img.mode != "RGB":
                    img = img.convert("RGB")

                img = ImageOps.fit(
                    img,
                    (1200, 675),
                    method=Image.Resampling.LANCZOS,
                )

                img_buf = io.BytesIO()

                img.save(
                    img_buf,
                    format="JPEG",
                    quality=92,
                )

                img_buf.seek(0)

                xl_img = OpenPyXLImage(img_buf)

                xl_img.width = 440
                xl_img.height = 250

                ws.add_image(
                    xl_img,
                    f"{col_letter}{photo_row}",
                )

            curr_row += 1

        curr_row += 1

    # --------------------------------------------------------------------------
    # 5. INSPECTION CONCLUSION
    # --------------------------------------------------------------------------

    ws.merge_cells(
        start_row=curr_row,
        start_column=2,
        end_row=curr_row,
        end_column=9,
    )

    ws.row_dimensions[curr_row].height = 22

    conc_cell = ws.cell(
        row=curr_row,
        column=2,
        value="INSPECTION COMMENTS & CONCLUSION",
    )

    conc_cell.fill = PatternFill(
        start_color=c_header,
        fill_type="solid",
    )

    conc_cell.font = Font(
        name="Calibri",
        size=10,
        bold=True,
        color="FFFFFF",
    )

    conc_cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
        indent=1,
    )

    apply_range_style(
        ws,
        2, curr_row, 9, curr_row,
        border=thin_border,
    )

    curr_row += 1

    ws.merge_cells(
        start_row=curr_row,
        start_column=2,
        end_row=curr_row + 3,
        end_column=9,
    )

    for r in range(curr_row, curr_row + 4):
        ws.row_dimensions[r].height = 18

    text_cell = ws.cell(
        row=curr_row,
        column=2,
        value=report_data.get(
            "comments",
            "Inspection carried out in accordance with standard specifications.",
        ),
    )

    text_cell.alignment = Alignment(
        horizontal="left",
        vertical="top",
        wrap_text=True,
    )

    text_cell.font = Font(
        name="Calibri",
        size=9,
        color="0F172A",
    )

    apply_range_style(
        ws,
        2, curr_row, 9, curr_row + 3,
        border=thin_border,
    )

    # --------------------------------------------------------------------------
    # RETURN EXCEL BYTES
    # --------------------------------------------------------------------------

    out = io.BytesIO()

    wb.save(out)

    return out.getvalue()


# ==============================================================================
# STREAMLIT CONFIGURATION
# ==============================================================================

st.set_page_config(
    page_title="QC Inspector Cloud",
    page_icon="📋",
    layout="wide",
)


# ==============================================================================
# SUPABASE CONNECTION
# ==============================================================================

@st.cache_resource
def get_supabase() -> Client:

    supabase_url = st.secrets.get("SUPABASE_URL", "")
    supabase_key = st.secrets.get("SUPABASE_KEY", "")

    if not supabase_url or not supabase_key:
        raise RuntimeError(
            "SUPABASE_URL или SUPABASE_KEY отсутствует "
            "в настройках Secrets."
        )

    supabase_url = str(supabase_url).strip()
    supabase_key = str(supabase_key).strip()

    if not supabase_url.startswith("https://"):
        raise RuntimeError(
            "SUPABASE_URL должен начинаться с https://"
        )

    return create_client(
        supabase_url,
        supabase_key,
    )


try:

    supabase = get_supabase()

except Exception as e:

    st.error(
        "Не удалось подключиться к Supabase. "
        "Проверь настройки Secrets."
    )

    st.code(str(e))

    st.stop()


# ==============================================================================
# SUPABASE STORAGE AND DATABASE HELPERS
# ==============================================================================

def _response_data_list(res) -> list:
    """Normalize Supabase response data into a list of dictionaries."""

    data = getattr(res, "data", None)

    if data is None and isinstance(res, dict):
        data = res.get("data")

    if isinstance(data, dict):
        return [data]

    if isinstance(data, list):
        return data

    return []


def _storage_object_path(excel_url) -> str | None:
    """
    Return a safe object path within qc-reports.

    The database stores a Storage path, not a signed URL.
    """

    if excel_url is None:
        return None

    value = str(excel_url).strip()

    if not value:
        return None

    if "://" in value:
        return None

    prefix = f"{QC_REPORTS_BUCKET}/"

    if value.startswith(prefix):
        value = value[len(prefix):]

    value = value.lstrip("/")

    parts = value.split("/")

    if not value or any(part in ("", ".", "..") for part in parts):
        return None

    return value


def _excel_url_filled(rep: dict) -> bool:

    value = rep.get("excel_url")

    if value is None:
        return False

    return bool(str(value).strip())


def _safe_download_name(
    report_no: str,
    result_status: str,
) -> str:
    """Create a safe filename for the Excel report."""

    raw = f"{report_no}_{result_status}.xlsx"

    for char in ("/", "\\", "\x00", ":", "*", "?", '"', "<", ">", "|"):
        raw = raw.replace(char, "_")

    return raw


def _extract_signed_url(res) -> str | None:
    """Extract signed URL from different Supabase response formats."""

    if isinstance(res, dict):

        data = res.get("data")

        if isinstance(data, dict):

            return (
                data.get("signedURL")
                or data.get("signedUrl")
                or data.get("signed_url")
            )

        return (
            res.get("signedURL")
            or res.get("signedUrl")
            or res.get("signed_url")
        )

    data = getattr(res, "data", None)

    if isinstance(data, dict):

        return (
            data.get("signedURL")
            or data.get("signedUrl")
            or data.get("signed_url")
        )

    return (
        getattr(res, "signedURL", None)
        or getattr(res, "signedUrl", None)
    )


def save_report_with_excel(
    row_fields: dict,
    excel_bytes: bytes,
) -> tuple[bool, str | None]:
    """
    Save a report in three stages:

    1. Insert the report into the reports table.
    2. Upload the Excel file to Supabase Storage.
    3. Save the Storage path to reports.excel_url and verify it.

    Returns:
        (success, error_message)
    """

    user = st.session_state.get("auth_user")

    if not user:
        return False, "Пользователь не авторизован."

    if not excel_bytes:
        return False, "Excel-файл пуст."

    user_id = str(user.id)

    report_id = None
    storage_path = None

    # --------------------------------------------------------------------------
    # STAGE 1. INSERT DATABASE RECORD
    # --------------------------------------------------------------------------

    try:

        insert_res = (
            supabase.table("reports")
            .insert(row_fields)
            .select("id")
            .execute()
        )

        rows = _response_data_list(insert_res)

        if not rows or not rows[0].get("id"):

            return (
                False,
                "Supabase не вернул ID созданного отчёта. "
                "Проверь таблицу reports и права INSERT/SELECT.",
            )

        report_id = str(rows[0]["id"])

    except Exception as e:

        return (
            False,
            f"Не удалось создать запись в reports: {e}",
        )

    # --------------------------------------------------------------------------
    # STAGE 2. UPLOAD EXCEL TO STORAGE
    # --------------------------------------------------------------------------

    storage_path = f"{user_id}/{report_id}/report.xlsx"

    try:

        supabase.storage.from_(
            QC_REPORTS_BUCKET
        ).upload(
            storage_path,
            excel_bytes,
            file_options={
                "content-type": EXCEL_MIME,
                "upsert": "false",
            },
        )

    except Exception as e:

        return (
            False,
            f"Запись отчёта создана, но Excel не загрузился "
            f"в Storage: {e}",
        )

    # --------------------------------------------------------------------------
    # STAGE 3. SAVE STORAGE PATH TO DATABASE
    # --------------------------------------------------------------------------

    try:

        (
            supabase.table("reports")
            .update({
                "excel_url": storage_path,
            })
            .eq("id", report_id)
            .eq("inspector_id", user_id)
            .execute()
        )

    except Exception as e:

        return (
            False,
            f"Excel загружен в Storage, но путь не удалось "
            f"сохранить в reports.excel_url: {e}",
        )

    # --------------------------------------------------------------------------
    # STAGE 4. VERIFY DATABASE VALUE
    # --------------------------------------------------------------------------

    try:

        check_res = (
            supabase.table("reports")
            .select("excel_url")
            .eq("id", report_id)
            .eq("inspector_id", user_id)
            .limit(1)
            .execute()
        )

        checked = _response_data_list(check_res)

        if not checked:

            return (
                False,
                "Excel загружен, но запись не удалось "
                "прочитать для проверки.",
            )

        saved_path = checked[0].get("excel_url")

        if saved_path != storage_path:

            return (
                False,
                "Excel загружен, но путь в базе не совпадает "
                "с ожидаемым. Проверь UPDATE policy для reports.",
            )

    except Exception as e:

        return (
            False,
            f"Не удалось проверить сохранённый путь Excel: {e}",
        )

    return True, None


def download_excel_from_storage(
    storage_path: str,
) -> tuple[bytes | None, str | None]:
    """
    Download Excel from Supabase Storage.

    First tries the Storage SDK.
    If that fails, tries a temporary signed URL.
    """

    storage_path = _storage_object_path(storage_path)

    if not storage_path:

        return None, "Некорректный путь к файлу."

    # --------------------------------------------------------------------------
    # METHOD 1. STORAGE SDK
    # --------------------------------------------------------------------------

    try:

        data = supabase.storage.from_(
            QC_REPORTS_BUCKET
        ).download(storage_path)

        if isinstance(data, (bytes, bytearray)) and data:

            return bytes(data), None

    except Exception as e:

        sdk_error = str(e)

    else:

        sdk_error = "Storage SDK не вернул содержимое файла."

    # --------------------------------------------------------------------------
    # METHOD 2. SIGNED URL
    # --------------------------------------------------------------------------

    try:

        signed_res = (
            supabase.storage.from_(QC_REPORTS_BUCKET)
            .create_signed_url(storage_path, 60)
        )

        signed_url = _extract_signed_url(signed_res)

    except Exception as e:

        return (
            None,
            f"Не удалось создать временную ссылку. "
            f"Storage SDK: {sdk_error}. Signed URL: {e}",
        )

    if not signed_url or "://" not in str(signed_url):

        return (
            None,
            f"Supabase не вернул корректную временную ссылку. "
            f"Storage SDK: {sdk_error}",
        )

    # --------------------------------------------------------------------------
    # DOWNLOAD USING SIGNED URL
    # --------------------------------------------------------------------------

    try:

        request = urllib.request.Request(
            signed_url,
            headers={
                "User-Agent": "QC-Inspector-Cloud/1.0",
            },
        )

        with urllib.request.urlopen(
            request,
            timeout=60,
        ) as response:

            body = response.read()

        if not body:

            return None, "Получен пустой файл."

        return body, None

    except Exception as e:

        return (
            None,
            f"Не удалось скачать Excel. "
            f"Storage SDK: {sdk_error}. HTTP: {e}",
        )


# ==============================================================================
# SESSION STATE
# ==============================================================================

if "auth_user" not in st.session_state:
    st.session_state.auth_user = None

if "reference_spec" not in st.session_state:
    st.session_state.reference_spec = None

if "measured_samples" not in st.session_state:
    st.session_state.measured_samples = []

if "comments_val" not in st.session_state:
    st.session_state.comments_val = ""

if "conclusion_val" not in st.session_state:
    st.session_state.conclusion_val = ""

if "last_excel_bytes" not in st.session_state:
    st.session_state.last_excel_bytes = None

if "last_excel_name" not in st.session_state:
    st.session_state.last_excel_name = "qc_report.xlsx"


# ==============================================================================
# AUTHENTICATION
# ==============================================================================

if not st.session_state.auth_user:

    st.markdown(
        """
        <h2 style="text-align:center;">
            QC Surveyor Portal
        </h2>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=True):

        auth_mode = st.radio(
            "Режим входа",
            ["Вход", "Регистрация"],
            horizontal=True,
        )

        email = st.text_input(
            "Email",
        )

        pwd = st.text_input(
            "Пароль",
            type="password",
        )

        if auth_mode == "Вход":

            if st.button(
                "Войти",
                type="primary",
                use_container_width=True,
            ):

                if not email.strip() or not pwd:

                    st.warning(
                        "Введите email и пароль."
                    )

                else:

                    try:

                        res = supabase.auth.sign_in_with_password({
                            "email": email.strip(),
                            "password": pwd,
                        })

                        if res.user:

                            st.session_state.auth_user = res.user

                            st.rerun()

                        else:

                            st.error(
                                "Не удалось авторизовать пользователя."
                            )

                    except Exception as e:

                        st.error(
                            "Не удалось войти. Проверьте email и пароль."
                        )

                        st.caption(str(e))

        elif auth_mode == "Регистрация":

            if st.button(
                "Создать аккаунт",
                type="primary",
                use_container_width=True,
            ):

                if not email.strip() or not pwd:

                    st.warning(
                        "Введите email и пароль."
                    )

                elif len(pwd) < 6:

                    st.warning(
                        "Используйте пароль длиной не менее 6 символов."
                    )

                else:

                    try:

                        res = supabase.auth.sign_up({
                            "email": email.strip(),
                            "password": pwd,
                        })

                        if res.user:

                            if res.session:

                                st.session_state.auth_user = res.user

                                st.success(
                                    "Аккаунт создан. Вы вошли в систему."
                                )

                                st.rerun()

                            else:

                                st.success(
                                    "Регистрация отправлена. "
                                    "Если в Supabase включено подтверждение "
                                    "email, подтвердите адрес и затем войдите."
                                )

                        else:

                            st.error(
                                "Supabase не вернул данные нового пользователя."
                            )

                    except Exception as e:

                        st.error(
                            "Не удалось создать аккаунт."
                        )

                        st.caption(str(e))

    st.stop()


# ==============================================================================
# USER HEADER
# ==============================================================================

c_user, c_out = st.columns([5, 1])

c_user.write(
    f"Инспектор: **{st.session_state.auth_user.email}**"
)

if c_out.button(
    "Выйти",
    use_container_width=True,
):

    try:
        supabase.auth.sign_out()
    except Exception:
        pass

    st.session_state.auth_user = None

    st.session_state.last_excel_bytes = None

    st.session_state.last_excel_name = "qc_report.xlsx"

    st.rerun()


# ==============================================================================
# MAIN TABS
# ==============================================================================

tab_new, tab_archive = st.tabs([
    "📝 Новый отчет",
    "🗄️ Архив отчетов",
])


# ==============================================================================
# TAB 1. REPORT ARCHIVE
# ==============================================================================

with tab_archive:

    st.subheader(
        "Сохраненные инспекции"
    )

    try:

        reports = (
            supabase.table("reports")
            .select("*")
            .eq(
                "inspector_id",
                st.session_state.auth_user.id,
            )
            .order(
                "created_at",
                desc=True,
            )
            .execute()
        )

        report_list = reports.data or []

        if report_list:

            for rep in report_list:

                with st.container(border=True):

                    c1, c2, c3, c4 = st.columns(
                        [2, 3, 2, 2]
                    )

                    c1.write(
                        f"**{rep.get('report_no', '-')}** "
                        f"({rep.get('inspection_date') or '-'})"
                    )

                    c2.write(
                        f"Товар: {rep.get('item_name') or '-'} "
                        f"| Поставщик: "
                        f"{rep.get('supplier_name') or '-'}"
                    )

                    c3.write(
                        f"Статус: "
                        f"**{rep.get('result_status') or '-'}**"
                    )

                    rid = str(
                        rep.get("id") or ""
                    )

                    storage_path = _storage_object_path(
                        rep.get("excel_url")
                    )

                    if rid and storage_path:

                        if c4.button(
                            "Скачать Excel",
                            key=f"btn_arch_{rid}",
                            use_container_width=True,
                        ):

                            with st.spinner(
                                "Загрузка Excel из облака..."
                            ):

                                file_bytes, download_error = (
                                    download_excel_from_storage(
                                        storage_path
                                    )
                                )

                            if file_bytes:

                                st.session_state[
                                    f"arch_excel_{rid}"
                                ] = file_bytes

                                st.session_state[
                                    f"arch_name_{rid}"
                                ] = _safe_download_name(
                                    str(
                                        rep.get("report_no")
                                        or "report"
                                    ),
                                    str(
                                        rep.get("result_status")
                                        or ""
                                    ),
                                )

                            else:

                                st.warning(
                                    "Не удалось получить Excel "
                                    "из облачного хранилища."
                                )

                                if download_error:
                                    st.caption(download_error)

                        cached = st.session_state.get(
                            f"arch_excel_{rid}"
                        )

                        if cached:

                            st.download_button(
                                label="📥 Сохранить файл",
                                data=cached,
                                file_name=st.session_state.get(
                                    f"arch_name_{rid}",
                                    "qc_report.xlsx",
                                ),
                                mime=EXCEL_MIME,
                                key=f"dl_arch_{rid}",
                                use_container_width=True,
                            )

                    else:

                        c4.caption(
                            "Excel ещё не сохранён в архиве"
                        )

        else:

            st.info(
                "История отчетов пока пуста."
            )

    except Exception as e:

        st.warning(
            "Не удалось загрузить архив. "
            "Проверьте подключение к Supabase и права доступа."
        )

        st.caption(str(e))


# ==============================================================================
# TAB 2. NEW INSPECTION REPORT
# ==============================================================================

with tab_new:

    st.subheader(
        "Конфигурация инспекции"
    )

    # ==========================================================================
    # 1. BATCH AND SUPPLIER INFORMATION
    # ==========================================================================

    with st.container(border=True):

        st.markdown(
            "**1. Данные партии и поставщика**"
        )

        batch_source = st.radio(
            "Источник данных",
            [
                "Ввести вручную",
                "Распознать стикер/коробку по фото",
            ],
            horizontal=True,
        )

        extracted_data = {}

        if batch_source == "Распознать стикер/коробку по фото":

            sticker_file = st.file_uploader(
                "Фото шиппинг-марки/стикера",
                type=["jpg", "png", "jpeg"],
                key="sticker_upload",
            )

            if sticker_file and st.button(
                "⚡ Распознать через ИИ",
                key="btn_extract_batch",
            ):

                with st.spinner(
                    "Извлечение данных через ИИ..."
                ):

                    try:

                        extracted_data = (
                            ai_service.extract_batch_info_from_image(
                                sticker_file.getvalue()
                            )
                        ) or {}

                        st.session_state[
                            "extracted_batch_data"
                        ] = extracted_data

                        st.rerun()

                    except Exception as e:

                        st.error(
                            "Не удалось распознать данные."
                        )

                        st.caption(str(e))

            extracted_data = st.session_state.get(
                "extracted_batch_data",
                {},
            )

        col_b1, col_b2, col_b3 = st.columns(3)

        report_no = col_b1.text_input(
            "Report No.",
            value=(
                "QC-"
                + st.session_state.auth_user.id[:5].upper()
            ),
        )

        client_name = col_b2.text_input(
            "Клиент / Заказчик"
        )

        supplier_name = col_b3.text_input(
            "Фабрика / Поставщик",
            value=extracted_data.get(
                "supplier",
                "",
            ),
        )

        col_b4, col_b5, col_b6 = st.columns(3)

        order_no = col_b4.text_input(
            "PO / Order No.",
            value=extracted_data.get(
                "order_no",
                "",
            ),
        )

        item_name = col_b5.text_input(
            "Наименование товара",
            value=extracted_data.get(
                "item_name",
                "",
            ),
        )

        sku = col_b6.text_input(
            "SKU / Артикул",
            value=extracted_data.get(
                "sku",
                "",
            ),
        )

    # ==========================================================================
    # 2. GARMENT MEASUREMENT CHART
    # ==========================================================================

    with st.container(border=True):

        st.markdown(
            "**2. Размерная сетка "
            "(Garment Measurement Chart)**"
        )

        use_spec = st.checkbox(
            "Включить модуль размерной сетки для этой инспекции",
            value=True,
        )

        if use_spec:

            if not st.session_state.reference_spec:

                st.caption(
                    "Загрузите 1 или 2 фото листа спецификации из техпака."
                )

                spec_imgs = st.file_uploader(
                    "Фото эталонной спецификации (до 2 фото)",
                    type=["jpg", "png", "jpeg"],
                    accept_multiple_files=True,
                    key="spec_upload",
                )

                if spec_imgs:

                    if len(spec_imgs) > 2:

                        st.warning(
                            "Загружено больше 2 фото. "
                            "Первые 2 будут взяты в обработку."
                        )

                        spec_imgs = spec_imgs[:2]

                    cols = st.columns(
                        len(spec_imgs)
                    )

                    for idx, img_file in enumerate(spec_imgs):

                        with cols[idx]:

                            st.image(
                                img_file,
                                caption=f"Фото эталона #{idx + 1}",
                                use_container_width=True,
                            )

                    if st.button(
                        "⚡ Оцифровать эталон (объединить фото)",
                        type="primary",
                        use_container_width=True,
                    ):

                        with st.spinner(
                            "Нейросеть считывает и объединяет данные с фото..."
                        ):

                            try:

                                bytes_list = [
                                    f.getvalue()
                                    for f in spec_imgs
                                ]

                                spec = (
                                    ai_service.digitize_spec_sheet(
                                        bytes_list
                                    )
                                )

                                if (
                                    spec
                                    and isinstance(spec, dict)
                                    and "poms" in spec
                                ):

                                    st.session_state.reference_spec = spec

                                    st.session_state.measured_samples = []

                                    st.rerun()

                                else:

                                    st.error(
                                        "Не удалось извлечь таблицу. "
                                        "Проверьте чёткость фото."
                                    )

                            except Exception as e:

                                st.error(
                                    "Ошибка при оцифровке спецификации."
                                )

                                st.caption(str(e))

            else:

                spec = st.session_state.reference_spec

                st.success(
                    "Эталон загружен. Размеры: "
                    + ", ".join(
                        str(x)
                        for x in spec.get("sizes", [])
                    )
                    + f" ({spec.get('unit', 'cm')})"
                )

                if st.button(
                    "Сбросить эталонную сетку"
                ):

                    st.session_state.reference_spec = None

                    st.session_state.measured_samples = []

                    st.rerun()

                sizes = spec.get("sizes", [])

                if sizes:

                    s_num = (
                        len(st.session_state.measured_samples)
                        + 1
                    )

                    st.markdown(
                        f"##### ➕ Замер образца #{s_num}"
                    )

                    selected_size = st.selectbox(
                        "Размер проверяемого образца",
                        sizes,
                        key=f"sz_{s_num}",
                    )

                    inputs = {}

                    with st.form(
                        key=f"sample_form_{s_num}"
                    ):

                        for pom in spec.get("poms", []):

                            p_code = pom.get(
                                "code",
                                "",
                            )

                            p_desc = pom.get(
                                "description",
                                "",
                            )

                            ideal_val = (
                                pom.get("specs", {}).get(
                                    selected_size,
                                    0.0,
                                )
                            )

                            t_min = float(
                                pom.get("tol_minus", 1.0)
                            )

                            t_plus = float(
                                pom.get("tol_plus", 1.0)
                            )

                            c1, c2, c3 = st.columns(
                                [1, 4, 3]
                            )

                            c1.write(
                                f"**{p_code}**"
                            )

                            c2.write(
                                f"{p_desc} "
                                f"(−{t_min} / +{t_plus})"
                            )

                            inputs[p_code] = {
                                "val": c3.text_input(
                                    f"Замер {p_code}",
                                    placeholder=str(ideal_val),
                                    label_visibility="collapsed",
                                    key=f"measurement_{s_num}_{p_code}",
                                ),
                                "ideal": ideal_val,
                                "desc": p_desc,
                                "t_min": t_min,
                                "t_plus": t_plus,
                            }

                        if st.form_submit_button(
                            "✓ Проверить образец"
                        ):

                            sample_items = []

                            validation_errors = []

                            for code, item_meta in inputs.items():

                                raw = (
                                    item_meta["val"]
                                    .strip()
                                    .replace(",", ".")
                                )

                                try:

                                    if raw:
                                        act = float(raw)
                                    else:
                                        act = float(
                                            item_meta["ideal"]
                                        )

                                except (ValueError, TypeError):

                                    validation_errors.append(
                                        f"Некорректный замер: {code}"
                                    )

                                    continue

                                ideal = float(
                                    item_meta["ideal"]
                                )

                                diff = round(
                                    act - ideal,
                                    2,
                                )

                                is_fail = (
                                    diff < -item_meta["t_min"]
                                    or diff > item_meta["t_plus"]
                                )

                                sample_items.append({
                                    "code": code,
                                    "desc": item_meta["desc"],
                                    "actual": act,
                                    "diff": diff,
                                    "status": (
                                        "FAIL"
                                        if is_fail
                                        else "PASS"
                                    ),
                                })

                            if validation_errors:

                                for error in validation_errors:
                                    st.error(error)

                            elif sample_items:

                                st.session_state.measured_samples.append({
                                    "sample_no": s_num,
                                    "size": selected_size,
                                    "items": sample_items,
                                })

                                st.rerun()

                    if st.session_state.measured_samples:

                        st.markdown(
                            "**Готовые замеры в отчете:**"
                        )

                        for smp in st.session_state.measured_samples:

                            fails = sum(
                                1
                                for it in smp["items"]
                                if it["status"] == "FAIL"
                            )

                            st.write(
                                f"- Образец #{smp['sample_no']} "
                                f"({smp['size']}): "
                                + (
                                    f"Отклонений: {fails}"
                                    if fails
                                    else "В норме"
                                )
                            )

                else:

                    st.warning(
                        "В спецификации нет доступных размеров."
                    )

    # ==========================================================================
    # 3. INSPECTION PHOTOS
    # ==========================================================================

    photos_data = {
        "products": [],
        "packing": [],
        "labels": [],
        "defects": [],
    }

    with st.container(border=True):

        st.markdown(
            "**3. Фотоматериалы инспекции "
            "(до 12 фото на секцию)**"
        )

        sec_names = {
            "products": "Продукция (Products)",
            "packing": "Упаковка и маркировка (Packaging)",
            "labels": "Стикеры и штрихкоды (Labels)",
            "defects": "Дефекты (Defects)",
        }

        for sec_key, sec_title in sec_names.items():

            st.markdown(
                f"##### {sec_title}"
            )

            files = st.file_uploader(
                f"Загрузить фото для секции '{sec_title}'",
                type=["jpg", "png", "jpeg"],
                accept_multiple_files=True,
                key=f"f_{sec_key}",
                label_visibility="collapsed",
            )

            if files:

                if len(files) > 12:

                    st.warning(
                        f"Выбрано {len(files)} фото. "
                        "В отчёт попадут первые 12."
                    )

                    files = files[:12]

                preview_cols = st.columns(4)

                for idx, fl in enumerate(files):

                    with preview_cols[idx % 4]:

                        st.image(
                            fl,
                            use_container_width=True,
                        )

                    photos_data[sec_key].append({
                        "bytes": fl.getvalue(),
                    })

            st.divider()

    # ==========================================================================
    # 4. INSPECTOR COMMENTS
    # ==========================================================================

    with st.container(border=True):

        st.markdown(
            "**4. Комментарии**"
        )

        raw_comments = st.text_area(
            "Комментарии",
            value=st.session_state.comments_val,
            placeholder="Рабочие заметки инспектора...",
            label_visibility="collapsed",
            key="raw_comments_input",
        )

        if st.button(
            "✨ Рецензировать с ИИ",
            key="btn_ai_comments",
        ):

            if raw_comments.strip():

                with st.spinner(
                    "Рецензирование..."
                ):

                    try:

                        polished = ai_service.polish_qc_text(
                            raw_comments,
                            field_type="comments",
                        )

                        st.session_state.comments_val = polished

                        st.rerun()

                    except Exception as e:

                        st.error(
                            "Не удалось обработать комментарии."
                        )

                        st.caption(str(e))

    # ==========================================================================
    # 5. FINAL INSPECTION RESULT
    # ==========================================================================

    with st.container(border=True):

        st.markdown(
            "**5. Общий итог проверки**"
        )

        default_conclusion = (
            st.session_state.conclusion_val
            or st.session_state.comments_val
        )

        final_comments = st.text_area(
            "Итоговые выводы для отчета",
            value=default_conclusion,
            placeholder="Итоговое заключение для отчёта...",
            label_visibility="collapsed",
            key="final_comments_input",
        )

        if st.button(
            "✨ Рецензировать с ИИ",
            key="btn_ai_conclusion",
        ):

            if final_comments.strip():

                with st.spinner(
                    "Рецензирование..."
                ):

                    try:

                        polished = ai_service.polish_qc_text(
                            final_comments,
                            field_type="conclusion",
                        )

                        st.session_state.conclusion_val = polished

                        st.rerun()

                    except Exception as e:

                        st.error(
                            "Не удалось обработать заключение."
                        )

                        st.caption(str(e))

        result_status = st.selectbox(
            "Финальный результат инспекции",
            [
                "PASSED",
                "PENDING",
                "REJECTED",
            ],
        )

    # ==========================================================================
    # 6. GENERATE AND SAVE REPORT
    # ==========================================================================

    st.divider()

    if st.button(
        "🚀 Сформировать и сохранить отчет Excel",
        type="primary",
        use_container_width=True,
    ):

        if not report_no.strip():

            st.error(
                "Укажи номер отчёта."
            )

        else:

            rep_payload = {
                "report_no": report_no.strip(),
                "company_name": st.secrets.get(
                    "DEFAULT_COMPANY_NAME",
                    "BIZNESKIT",
                ),
                "client_name": client_name,
                "supplier_name": supplier_name,
                "order_no": order_no,
                "item_name": item_name,
                "sku": sku,
                "inspector_id": (
                    st.session_state.auth_user.email
                ),
                "inspection_date": date.today().isoformat(),
                "result_status": result_status,
                "comments": final_comments,
            }

            try:

                # --------------------------------------------------------------
                # GENERATE EXCEL
                # --------------------------------------------------------------

                with st.spinner(
                    "Генерация официального Excel-отчёта..."
                ):

                    excel_bytes = create_full_qc_report(
                        report_data=rep_payload,
                        photos_by_section=photos_data,
                        spec_data=(
                            st.session_state.reference_spec
                            if use_spec
                            else None
                        ),
                        measured_samples=(
                            st.session_state.measured_samples
                            if use_spec
                            else []
                        ),
                    )

                if not excel_bytes:

                    st.error(
                        "Не удалось создать Excel-файл."
                    )

                    st.stop()

                # --------------------------------------------------------------
                # CACHE FILE FOR DOWNLOAD IN CURRENT SESSION
                # --------------------------------------------------------------

                safe_name = _safe_download_name(
                    report_no.strip(),
                    result_status,
                )

                st.session_state.last_excel_bytes = excel_bytes

                st.session_state.last_excel_name = safe_name

                # --------------------------------------------------------------
                # PREPARE DATABASE RECORD
                # --------------------------------------------------------------

                row_fields = {
                    "inspector_id": (
                        st.session_state.auth_user.id
                    ),
                    "report_no": report_no.strip(),
                    "client_name": client_name,
                    "supplier_name": supplier_name,
                    "order_no": order_no,
                    "item_name": item_name,
                    "sku": sku,
                    "inspection_date": date.today().isoformat(),
                    "result_status": result_status,
                    "report_data": rep_payload,
                }

                # --------------------------------------------------------------
                # SAVE REPORT AND EXCEL TO CLOUD
                # --------------------------------------------------------------

                with st.spinner(
                    "Сохранение отчёта в Supabase..."
                ):

                    cloud_saved, cloud_error = (
                        save_report_with_excel(
                            row_fields=row_fields,
                            excel_bytes=excel_bytes,
                        )
                    )

                # --------------------------------------------------------------
                # SHOW RESULT
                # --------------------------------------------------------------

                if cloud_saved:

                    st.success(
                        "Отчёт успешно сформирован "
                        "и сохранён в облачный архив."
                    )

                else:

                    st.warning(
                        "Excel сформирован, но облачное "
                        "сохранение не завершено."
                    )

                    if cloud_error:
                        st.code(cloud_error)

                    st.info(
                        "Файл можно скачать ниже. "
                        "В архиве он может отсутствовать "
                        "или быть недоступен, пока проблема "
                        "с сохранением не будет устранена."
                    )

            except Exception as e:

                st.error(
                    "Не удалось сформировать Excel-файл."
                )

                st.caption(
                    str(e)
                )

    # ==========================================================================
    # 7. DOWNLOAD LAST GENERATED EXCEL
    # ==========================================================================

    if st.session_state.last_excel_bytes:

        st.download_button(
            label="📥 Скачать QC Report (.xlsx)",
            data=st.session_state.last_excel_bytes,
            file_name=st.session_state.last_excel_name,
            mime=EXCEL_MIME,
            use_container_width=True,
            key="dl_last_excel",
        )

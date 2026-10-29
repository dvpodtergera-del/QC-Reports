from __future__ import annotations

import sys
import os
import io
import urllib.request
import streamlit as st
from supabase import create_client, Client
from PIL import Image, ImageOps
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenPyXLImage

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import ai_service

# ==============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ СТИЛИЗАЦИИ EXCEL
# ==============================================================================
def apply_range_style(ws, min_col, min_row, max_col, max_row, border=None, fill=None):
    """Применяет рамки и заливку ко всем ячейкам диапазона."""
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if border:
                cell.border = border
            if fill:
                cell.fill = fill

# ==============================================================================
# МОДУЛЬ ГЕНЕРАЦИИ EXCEL (БЕЗ ПОДПИСЕЙ, ЧИСТАЯ HD-ФОТОСЕТКА)
# ==============================================================================
def create_full_qc_report(
    report_data: dict, 
    photos_by_section: dict, 
    spec_data: dict = None, 
    measured_samples: list = None
) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QC Inspection Report"
    ws.views.sheetView[0].showGridLines = True

    c_brand = "1E293B"      # Dark Slate (Главная шапка)
    c_header = "334155"     # Slate 700 (Разделители)
    c_sub = "475569"        # Slate 600
    c_border = "CBD5E1"     # Slate 300 (Тонкие границы ячеек)
    c_meta_bg = "F8FAFC"    # Светлый фон плашек
    
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
        bottom=Side(style="thin", color=c_border)
    )

    # Модульная сетка: Колонки B..I по 15 пунктов = симметрия для 2 фотокарт
    ws.column_dimensions["A"].width = 3
    for col_letter in ["B", "C", "D", "E", "F", "G", "H", "I"]:
        ws.column_dimensions[col_letter].width = 15
    for col_letter in ["J", "K", "L", "M", "N", "O"]:
        ws.column_dimensions[col_letter].width = 14

    # 1. Главная брендовая плашка Bizneskit
    ws.merge_cells("B2:I2")
    ws.row_dimensions[2].height = 40
    company_title = report_data.get("company_name", "BIZNESKIT").upper()
    h_cell = ws["B2"]
    h_cell.value = f"{company_title} — QUALITY CONTROL INSPECTION REPORT"
    h_cell.fill = PatternFill(start_color=c_brand, end_color=c_brand, fill_type="solid")
    h_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    h_cell.alignment = Alignment(horizontal="center", vertical="center")
    apply_range_style(ws, 2, 2, 9, 2, border=thin_border)

    # 2. Информационный блок партии
    res_status = report_data.get("result_status", "PENDING").upper()
    meta_rows = [
        ("REPORT NO:", report_data.get("report_no", "QC-001"), "INSPECTION DATE:", str(report_data.get("inspection_date", ""))),
        ("CLIENT:", report_data.get("client_name", "-"), "INSPECTOR ID:", str(report_data.get("inspector_id", "-"))),
        ("SUPPLIER / FACTORY:", report_data.get("supplier_name", "-"), "PO / ORDER NO:", str(report_data.get("order_no", "-"))),
        ("ITEM / SKU:", f"{report_data.get('item_name', '-')} / {report_data.get('sku', '-')}", "OVERALL RESULT:", res_status)
    ]

    curr_row = 4
    for r_data in meta_rows:
        ws.row_dimensions[curr_row].height = 22
        
        c1 = ws.cell(row=curr_row, column=2, value=r_data[0])
        c1.fill = PatternFill(start_color=c_meta_bg, fill_type="solid")
        c1.font = Font(name="Calibri", size=9, bold=True, color=c_sub)
        c1.border = thin_border
        
        ws.merge_cells(start_row=curr_row, start_column=3, end_row=curr_row, end_column=5)
        c2 = ws.cell(row=curr_row, column=3, value=r_data[1])
        c2.font = Font(name="Calibri", size=9, bold=True, color="0F172A")
        apply_range_style(ws, 3, curr_row, 5, curr_row, border=thin_border)

        c3 = ws.cell(row=curr_row, column=6, value=r_data[2])
        c3.fill = PatternFill(start_color=c_meta_bg, fill_type="solid")
        c3.font = Font(name="Calibri", size=9, bold=True, color=c_sub)
        c3.border = thin_border

        ws.merge_cells(start_row=curr_row, start_column=7, end_row=curr_row, end_column=9)
        c4 = ws.cell(row=curr_row, column=7, value=r_data[3])
        apply_range_style(ws, 7, curr_row, 9, curr_row, border=thin_border)
        
        if r_data[2] == "OVERALL RESULT:":
            c4.alignment = Alignment(horizontal="center", vertical="center")
            if res_status == "PASSED":
                c4.fill = PatternFill(start_color=c_pass_bg, fill_type="solid")
                c4.font = Font(name="Calibri", size=11, bold=True, color=c_pass_fg)
            elif res_status == "REJECTED":
                c4.fill = PatternFill(start_color=c_fail_bg, fill_type="solid")
                c4.font = Font(name="Calibri", size=11, bold=True, color=c_fail_fg)
            else:
                c4.fill = PatternFill(start_color=c_pend_bg, fill_type="solid")
                c4.font = Font(name="Calibri", size=11, bold=True, color=c_pend_fg)
        else:
            c4.font = Font(name="Calibri", size=9, color="0F172A")
            c4.alignment = Alignment(horizontal="left", vertical="center")

        curr_row += 1

    curr_row += 1

    # 3. Размерная сетка (Spec Sheet)
    if spec_data and measured_samples:
        ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=9)
        ws.row_dimensions[curr_row].height = 24
        sec_cell = ws.cell(row=curr_row, column=2, value="GARMENT MEASUREMENT & TOLERANCE SPEC SHEET")
        sec_cell.fill = PatternFill(start_color=c_header, fill_type="solid")
        sec_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        sec_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        apply_range_style(ws, 2, curr_row, 9, curr_row, border=thin_border)
        curr_row += 1

        ws.row_dimensions[curr_row].height = 20
        spec_headers = [
            (2, 2, "Code"),
            (3, 5, "POM Description"),
            (6, 6, "Tol (-)"),
            (7, 7, "Tol (+)")
        ]
        for s_col, e_col, h_name in spec_headers:
            if s_col != e_col:
                ws.merge_cells(start_row=curr_row, start_column=s_col, end_row=curr_row, end_column=e_col)
            c = ws.cell(row=curr_row, column=s_col, value=h_name)
            c.fill = PatternFill(start_color=c_sub, fill_type="solid")
            c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="center", vertical="center")
            apply_range_style(ws, s_col, curr_row, e_col, curr_row, border=thin_border)

        sample_col = 8
        for sample in measured_samples:
            ws.cell(row=curr_row, column=sample_col, value=f"#{sample['sample_no']} ({sample['size']}) Act").font = Font(name="Calibri", size=8, bold=True, color="FFFFFF")
            ws.cell(row=curr_row, column=sample_col).fill = PatternFill(start_color=c_sub, fill_type="solid")
            ws.cell(row=curr_row, column=sample_col).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr_row, column=sample_col).border = thin_border

            ws.cell(row=curr_row, column=sample_col + 1, value="Diff").font = Font(name="Calibri", size=8, bold=True, color="FFFFFF")
            ws.cell(row=curr_row, column=sample_col + 1).fill = PatternFill(start_color=c_sub, fill_type="solid")
            ws.cell(row=curr_row, column=sample_col + 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr_row, column=sample_col + 1).border = thin_border
            sample_col += 2

        curr_row += 1

        for pom in spec_data.get("poms", []):
            ws.row_dimensions[curr_row].height = 18
            ws.cell(row=curr_row, column=2, value=pom.get("code", "")).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr_row, column=2).border = thin_border
            
            ws.merge_cells(start_row=curr_row, start_column=3, end_row=curr_row, end_column=5)
            ws.cell(row=curr_row, column=3, value=pom.get("description", "")).alignment = Alignment(horizontal="left", vertical="center")
            apply_range_style(ws, 3, curr_row, 5, curr_row, border=thin_border)

            ws.cell(row=curr_row, column=6, value=f"-{pom.get('tol_minus', 1.0)}").alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr_row, column=7, value=f"+{pom.get('tol_plus', 1.0)}").alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr_row, column=6).border = thin_border
            ws.cell(row=curr_row, column=7).border = thin_border

            s_col = 8
            for sample in measured_samples:
                item = next((x for x in sample.get("items", []) if x.get("code") == pom.get("code")), None)
                if item:
                    c_act = ws.cell(row=curr_row, column=s_col, value=item.get("actual", "-"))
                    c_diff = ws.cell(row=curr_row, column=s_col + 1, value=item.get("diff", 0))
                    
                    c_act.border = thin_border
                    c_diff.border = thin_border
                    c_act.alignment = Alignment(horizontal="center", vertical="center")
                    c_diff.alignment = Alignment(horizontal="center", vertical="center")
                    c_act.font = Font(name="Calibri", size=9)

                    if item.get("status") == "FAIL":
                        c_diff.fill = PatternFill(start_color=c_fail_bg, fill_type="solid")
                        c_diff.font = Font(name="Calibri", size=9, bold=True, color=c_fail_fg)
                    else:
                        c_diff.fill = PatternFill(start_color=c_pass_bg, fill_type="solid")
                        c_diff.font = Font(name="Calibri", size=9, bold=True, color=c_pass_fg)
                s_col += 2

            curr_row += 1
        curr_row += 1

    # 4. Фотоматериалы (HD 1200x675, БЕЗ подписей, чистое позиционирование)
    section_titles = {
        "products": "PRODUCT WORKMANSHIP PHOTOS",
        "packing": "PACKAGING & CARTON DETAILS",
        "labels": "LABELS & BARCODES VERIFICATION",
        "defects": "DEFECT DETAILS & LOG"
    }

    slots = [
        (2, 5, 'B'),  # Левая позиция: B..E
        (6, 9, 'F')   # Правая позиция: F..I
    ]

    for s_key, s_label in section_titles.items():
        photos = photos_by_section.get(s_key, [])
        if not photos:
            continue

        ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=9)
        ws.row_dimensions[curr_row].height = 22
        t_cell = ws.cell(row=curr_row, column=2, value=s_label)
        t_cell.fill = PatternFill(start_color=c_header, fill_type="solid")
        t_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        t_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        apply_range_style(ws, 2, curr_row, 9, curr_row, border=thin_border)
        curr_row += 1

        for i in range(0, len(photos), 2):
            photo_row = curr_row
            ws.row_dimensions[photo_row].height = 195

            for slot_idx, (sc, ec, col_letter) in enumerate(slots):
                p_idx = i + slot_idx
                if p_idx < len(photos):
                    p_info = photos[p_idx]
                    
                    ws.merge_cells(start_row=photo_row, start_column=sc, end_row=photo_row, end_column=ec)
                    apply_range_style(ws, sc, photo_row, ec, photo_row, border=thin_border)

                    img = Image.open(io.BytesIO(p_info["bytes"]))
                    img = ImageOps.exif_transpose(img)
                    
                    if img.mode in ("RGBA", "LA", "P"):
                        background = Image.new("RGB", img.size, (255, 255, 255))
                        if img.mode == "P":
                            img = img.convert("RGBA")
                        if "A" in img.mode:
                            background.paste(img, mask=img.split()[-1])
                            img = background
                        else:
                            img = img.convert("RGB")
                    elif img.mode != "RGB":
                        img = img.convert("RGB")

                    img = ImageOps.fit(img, (1200, 675), method=Image.Resampling.LANCZOS)
                    
                    img_buf = io.BytesIO()
                    img.save(img_buf, format="JPEG", quality=92)
                    img_buf.seek(0)

                    xl_img = OpenPyXLImage(img_buf)
                    xl_img.width = 440
                    xl_img.height = 250
                    ws.add_image(xl_img, f"{col_letter}{photo_row}")

            curr_row += 1
        curr_row += 1

    # 5. Общий итог проверки в Excel (Выводы и замечания)
    ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=9)
    ws.row_dimensions[curr_row].height = 22
    conc_cell = ws.cell(row=curr_row, column=2, value="INSPECTION COMMENTS & CONCLUSION")
    conc_cell.fill = PatternFill(start_color=c_header, fill_type="solid")
    conc_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    conc_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    apply_range_style(ws, 2, curr_row, 9, curr_row, border=thin_border)
    curr_row += 1

    ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row + 3, end_column=9)
    for r in range(curr_row, curr_row + 4):
        ws.row_dimensions[r].height = 18
    text_cell = ws.cell(row=curr_row, column=2, value=report_data.get("comments", "Inspection carried out in accordance with standard specifications."))
    text_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    text_cell.font = Font(name="Calibri", size=9, color="0F172A")
    apply_range_style(ws, 2, curr_row, 9, curr_row + 3, border=thin_border)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ==============================================================================
# ОСНОВНОЙ ИНТЕРФЕЙС STREAMLIT
# ==============================================================================
st.set_page_config(page_title="QC Inspector Cloud", page_icon="📋", layout="wide")

@st.cache_resource
def get_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

supabase = get_supabase()

QC_REPORTS_BUCKET = "qc-reports"
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _response_data_list(res) -> list:
    data = getattr(res, "data", None)
    if data is None and isinstance(res, dict):
        data = res.get("data")
    if isinstance(data, dict):
        return [data]
    if isinstance(data, list):
        return data
    return []


def _storage_call_failed(res) -> bool:
    if res is None:
        return True
    if isinstance(res, dict) and res.get("error"):
        return True
    error = getattr(res, "error", None)
    return bool(error)


def _safe_download_name(report_no: str, result_status: str) -> str:
    raw = f"{report_no}_{result_status}.xlsx"
    return raw.replace("/", "_").replace("\\", "_").replace("\x00", "")


def _storage_object_path(excel_url) -> str | None:
    """Возвращает путь объекта в bucket. Signed URL в БД не используем."""
    if excel_url is None:
        return None
    value = str(excel_url).strip()
    if not value or "://" in value:
        return None
    prefix = f"{QC_REPORTS_BUCKET}/"
    if value.startswith(prefix):
        value = value[len(prefix):]
    value = value.lstrip("/")
    if not value or ".." in value.split("/"):
        return None
    return value


def _excel_url_filled(rep: dict) -> bool:
    value = rep.get("excel_url")
    if value is None:
        return False
    return bool(str(value).strip())


def save_report_with_excel(row_fields: dict, excel_bytes: bytes) -> bool:
    """
    Сначала insert в reports, затем upload в Storage, затем путь в excel_url.
    True только если путь в записи подтверждён чтением из таблицы.
    """
    user_id = st.session_state.auth_user.id
    try:
        insert_res = supabase.table("reports").insert(row_fields).execute()
    except Exception:
        return False

    rows = _response_data_list(insert_res)
    if not rows or not rows[0].get("id"):
        return False

    report_id = str(rows[0]["id"])
    storage_path = f"{user_id}/{report_id}/report.xlsx"

    try:
        upload_res = supabase.storage.from_(QC_REPORTS_BUCKET).upload(
            storage_path,
            excel_bytes,
            file_options={
                "content-type": EXCEL_MIME,
                "upsert": "false",
            },
        )
        if _storage_call_failed(upload_res):
            return False
    except Exception:
        return False

    try:
        supabase.table("reports").update({"excel_url": storage_path}).eq("id", report_id).eq(
            "inspector_id", user_id
        ).execute()
        check_res = (
            supabase.table("reports")
            .select("excel_url")
            .eq("id", report_id)
            .eq("inspector_id", user_id)
            .limit(1)
            .execute()
        )
    except Exception:
        return False

    checked = _response_data_list(check_res)
    if not checked:
        return False
    return checked[0].get("excel_url") == storage_path


def download_excel_from_storage(storage_path: str) -> bytes | None:
    try:
        data = supabase.storage.from_(QC_REPORTS_BUCKET).download(storage_path)
    except Exception:
        data = None

    if isinstance(data, (bytes, bytearray)) and data:
        return bytes(data)

    try:
        signed = supabase.storage.from_(QC_REPORTS_BUCKET).create_signed_url(storage_path, 60)
    except Exception:
        return None

    signed_url = None
    if isinstance(signed, dict):
        signed_url = signed.get("signedURL") or signed.get("signedUrl")
    else:
        signed_url = getattr(signed, "signedURL", None) or getattr(signed, "signedUrl", None)
        if not signed_url and isinstance(getattr(signed, "data", None), dict):
            payload = signed.data
            signed_url = payload.get("signedURL") or payload.get("signedUrl")

    if not signed_url or "://" not in str(signed_url):
        return None

    try:
        with urllib.request.urlopen(signed_url, timeout=60) as resp:
            body = resp.read()
        return body if body else None
    except Exception:
        return None


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

# Авторизация
if not st.session_state.auth_user:
    st.markdown("<h2 style='text-align: center; color: #1E293B;'>QC Surveyor Portal</h2>", unsafe_allow_html=True)
    with st.container(border=True):
        auth_mode = st.radio("Режим входа", ["Вход", "Регистрация"], horizontal=True)
        email = st.text_input("Email")
        pwd = st.text_input("Пароль", type="password")
        
        if auth_mode == "Вход" and st.button("Войти", type="primary", use_container_width=True):
            try:
                res = supabase.auth.sign_in_with_password({"email": email, "password": pwd})
                if res.user:
                    st.session_state.auth_user = res.user
                    st.rerun()
            except Exception:
                st.error("Не удалось войти. Проверьте email и пароль.")

        elif auth_mode == "Регистрация" and st.button("Создать аккаунт", type="primary", use_container_width=True):
            try:
                res = supabase.auth.sign_up({"email": email, "password": pwd})
                st.success("Регистрация успешна. Теперь можно войти.")
            except Exception:
                st.error("Не удалось создать аккаунт. Попробуйте другой email или повторите позже.")
    st.stop()

# Панель пользователя
c_user, c_out = st.columns([5, 1])
c_user.write(f"Инспектор: **{st.session_state.auth_user.email}**")
if c_out.button("Выйти"):
    supabase.auth.sign_out()
    st.session_state.auth_user = None
    st.rerun()

tab_new, tab_archive = st.tabs(["📝 Новый отчет", "🗄️ Архив отчетов"])

with tab_archive:
    st.subheader("Сохраненные инспекции")

    try:
        reports = (
            supabase.table("reports")
            .select("*")
            .eq("inspector_id", st.session_state.auth_user.id)
            .order("created_at", desc=True)
            .execute()
        )

        report_list = reports.data or []

        if "selected_report_id" not in st.session_state:
            st.session_state.selected_report_id = None

        if report_list:
            for rep in report_list:
                rid = str(rep.get("id") or "")
                storage_path = _storage_object_path(rep.get("excel_url"))

                with st.container(border=True):
                    c1, c2, c3, c4, c5 = st.columns([2, 3, 2, 1.2, 1.2])

                    c1.write(f"**{rep.get('report_no') or '-'}**")
                    c1.caption(rep.get("inspection_date") or "-")

                    c2.write(f"Товар: {rep.get('item_name') or '-'}")
                    c2.caption(f"Поставщик: {rep.get('supplier_name') or '-'}")

                    c3.write(f"Статус: **{rep.get('result_status') or '-'}**")

                    if c4.button(
                        "Открыть",
                        key=f"open_report_{rid}",
                        use_container_width=True,
                        disabled=not rid,
                    ):
                        st.session_state.selected_report_id = rid
                        st.rerun()

                    if storage_path:
                        if c5.button(
                            "Скачать Excel",
                            key=f"btn_arch_{rid}",
                            use_container_width=True,
                        ):
                            with st.spinner("Загрузка Excel из облака..."):
                                file_bytes, download_error = download_excel_from_storage(storage_path)

                            if file_bytes:
                                st.session_state[f"arch_excel_{rid}"] = file_bytes
                                st.session_state[f"arch_name_{rid}"] = _safe_download_name(
                                    str(rep.get("report_no") or "report"),
                                    str(rep.get("result_status") or ""),
                                )
                            else:
                                st.warning("Не удалось получить Excel из облачного хранилища.")
                                if download_error:
                                    st.caption(download_error)

                        cached = st.session_state.get(f"arch_excel_{rid}")
                        if cached:
                            st.download_button(
                                label="📥 Сохранить файл",
                                data=cached,
                                file_name=st.session_state.get(f"arch_name_{rid}", "qc_report.xlsx"),
                                mime=EXCEL_MIME,
                                key=f"dl_arch_{rid}",
                                use_container_width=True,
                            )
                    else:
                        c5.caption("Excel не сохранен")

            # Просмотр выбранного отчета
            selected_id = st.session_state.get("selected_report_id")
            selected_report = next(
                (r for r in report_list if str(r.get("id") or "") == str(selected_id)),
                None,
            )

            if selected_report:
                st.divider()
                st.subheader(f"Просмотр отчета: {selected_report.get('report_no') or '-'}")

                if st.button("Закрыть отчет", key="close_report_view"):
                    st.session_state.selected_report_id = None
                    st.rerun()

                report_data = selected_report.get("report_data") or {}
                if isinstance(report_data, str):
                    import json
                    try:
                        report_data = json.loads(report_data)
                    except Exception:
                        report_data = {}

                if not isinstance(report_data, dict):
                    report_data = {}

                # Режим редактирования конкретного отчета
                edit_state_key = f"editing_report_{selected_id}"
                if edit_state_key not in st.session_state:
                    st.session_state[edit_state_key] = False

                edit_col, _ = st.columns([2, 5])
                if not st.session_state[edit_state_key]:
                    if edit_col.button(
                        "✏️ Редактировать отчет",
                        key=f"start_edit_{selected_id}",
                        use_container_width=True,
                    ):
                        st.session_state[edit_state_key] = True
                        st.rerun()

                if st.session_state[edit_state_key]:
                    st.markdown("### Редактирование отчета")
                    st.caption(
                        "Изменения сохраняются в базе данных. Уже созданный Excel-файл "
                        "останется прежним, пока не будет сформирован заново."
                    )

                    with st.form(key=f"edit_report_form_{selected_id}"):
                        edit_c1, edit_c2 = st.columns(2)
                        with edit_c1:
                            edited_report_no = st.text_input(
                                "Номер отчета",
                                value=str(selected_report.get("report_no") or ""),
                            )
                            edited_client = st.text_input(
                                "Клиент / Заказчик",
                                value=str(selected_report.get("client_name") or report_data.get("client_name") or ""),
                            )
                            edited_supplier = st.text_input(
                                "Фабрика / Поставщик",
                                value=str(selected_report.get("supplier_name") or report_data.get("supplier_name") or ""),
                            )
                            edited_order = st.text_input(
                                "Номер заказа",
                                value=str(selected_report.get("order_no") or report_data.get("order_no") or ""),
                            )
                            edited_item = st.text_input(
                                "Наименование товара",
                                value=str(selected_report.get("item_name") or report_data.get("item_name") or ""),
                            )
                        with edit_c2:
                            edited_sku = st.text_input(
                                "Артикул / SKU",
                                value=str(selected_report.get("sku") or report_data.get("sku") or ""),
                            )
                            edited_date = st.text_input(
                                "Дата инспекции (ГГГГ-ММ-ДД)",
                                value=str(selected_report.get("inspection_date") or report_data.get("inspection_date") or ""),
                            )
                            status_options = ["PASSED", "PENDING", "REJECTED"]
                            current_status = str(selected_report.get("result_status") or report_data.get("result_status") or "PENDING").upper()
                            if current_status not in status_options:
                                current_status = "PENDING"
                            edited_status = st.selectbox(
                                "Итог проверки",
                                status_options,
                                index=status_options.index(current_status),
                            )

                        edited_comments = st.text_area(
                            "Комментарии и заключение",
                            value=str(report_data.get("comments") or ""),
                            height=160,
                        )

                        save_col, cancel_col = st.columns(2)
                        save_changes = save_col.form_submit_button(
                            "💾 Сохранить изменения",
                            type="primary",
                            use_container_width=True,
                        )
                        cancel_changes = cancel_col.form_submit_button(
                            "Отмена",
                            use_container_width=True,
                        )

                    if cancel_changes:
                        st.session_state[edit_state_key] = False
                        st.rerun()

                    if save_changes:
                        from datetime import date as _date
                        try:
                            # Проверяем формат даты, если дата указана.
                            normalized_date = edited_date.strip()
                            if normalized_date:
                                _date.fromisoformat(normalized_date)

                            updated_report_data = dict(report_data)
                            updated_report_data.update({
                                "report_no": edited_report_no.strip(),
                                "client_name": edited_client.strip(),
                                "supplier_name": edited_supplier.strip(),
                                "order_no": edited_order.strip(),
                                "item_name": edited_item.strip(),
                                "sku": edited_sku.strip(),
                                "inspection_date": normalized_date,
                                "result_status": edited_status,
                                "comments": edited_comments,
                            })

                            update_fields = {
                                "report_no": edited_report_no.strip(),
                                "client_name": edited_client.strip(),
                                "supplier_name": edited_supplier.strip(),
                                "order_no": edited_order.strip(),
                                "item_name": edited_item.strip(),
                                "sku": edited_sku.strip(),
                                "inspection_date": normalized_date or None,
                                "result_status": edited_status,
                                "report_data": updated_report_data,
                            }

                            (
                                supabase.table("reports")
                                .update(update_fields)
                                .eq("id", selected_id)
                                .eq("inspector_id", st.session_state.auth_user.id)
                                .execute()
                            )

                            st.session_state[edit_state_key] = False
                            st.success("Изменения сохранены.")
                            st.rerun()

                        except ValueError:
                            st.error("Неверный формат даты. Используй ГГГГ-ММ-ДД, например 2026-10-29.")
                        except Exception as e:
                            st.error("Не удалось сохранить изменения. Проверь права UPDATE в Supabase.")
                            st.caption(str(e))

                # Показываем актуальные поля записи из таблицы, а также сохраненные детали.
                st.markdown("### Основная информация")
                left, right = st.columns(2)

                with left:
                    st.write("**Номер отчета:**", selected_report.get("report_no") or "-")
                    st.write("**Клиент:**", selected_report.get("client_name") or report_data.get("client_name") or "-")
                    st.write("**Фабрика / поставщик:**", selected_report.get("supplier_name") or report_data.get("supplier_name") or "-")
                    st.write("**Номер заказа:**", selected_report.get("order_no") or report_data.get("order_no") or "-")

                with right:
                    st.write("**Дата инспекции:**", selected_report.get("inspection_date") or report_data.get("inspection_date") or "-")
                    st.write("**Товар:**", selected_report.get("item_name") or report_data.get("item_name") or "-")
                    st.write("**Артикул:**", selected_report.get("sku") or report_data.get("sku") or "-")
                    st.write("**Итог проверки:**", selected_report.get("result_status") or report_data.get("result_status") or "-")

                st.markdown("### Комментарии и заключение")
                comments = report_data.get("comments") or ""
                if comments:
                    st.write(comments)
                else:
                    st.caption("Комментарии в сохраненных данных отсутствуют.")

                st.markdown("### Дополнительные данные")
                if report_data:
                    with st.expander("Показать все сохраненные данные отчета"):
                        st.json(report_data)
                else:
                    st.info("Для этой записи дополнительные данные отчета не сохранены.")

                # Фото находятся внутри Excel, если они были добавлены при его создании.
                st.caption(
                    "Примечание: просмотр показывает данные, сохраненные в базе. "
                    "Фотографии, встроенные в Excel, можно посмотреть, скачав Excel-файл."
                )

                selected_path = _storage_object_path(selected_report.get("excel_url"))
                if selected_path:
                    selected_cache_key = f"selected_excel_{selected_id}"

                    if st.button(
                        "Подготовить Excel для скачивания",
                        key=f"prepare_selected_excel_{selected_id}",
                    ):
                        with st.spinner("Загрузка Excel из облака..."):
                            file_bytes, download_error = download_excel_from_storage(selected_path)

                        if file_bytes:
                            st.session_state[selected_cache_key] = file_bytes
                        else:
                            st.error("Не удалось загрузить Excel.")
                            if download_error:
                                st.caption(download_error)

                    selected_bytes = st.session_state.get(selected_cache_key)
                    if selected_bytes:
                        st.download_button(
                            "📥 Скачать Excel этого отчета",
                            data=selected_bytes,
                            file_name=_safe_download_name(
                                str(selected_report.get("report_no") or "report"),
                                str(selected_report.get("result_status") or ""),
                            ),
                            mime=EXCEL_MIME,
                            key=f"download_selected_excel_{selected_id}",
                        )
                else:
                    st.info("Для этого отчета Excel-файл в облачном хранилище не найден.")

        else:
            st.info("История отчетов пока пуста.")

    except Exception as e:
        st.warning("Не удалось загрузить архив. Проверьте подключение к Supabase и права доступа.")
        st.caption(str(e))


with tab_new:
    st.subheader("Конфигурация инспекции")
    
    # 1. Реквизиты партии
    with st.container(border=True):
        st.markdown("**1. Данные партии и поставщика**")
        batch_source = st.radio("Источник данных", ["Ввести вручную", "Распознать стикер/коробку по фото"], horizontal=True)
        
        extracted_data = {}
        if batch_source == "Распознать стикер/коробку по фото":
            sticker_file = st.file_uploader("Фото шиппинг-марки/стикера", type=["jpg", "png", "jpeg"])
            if sticker_file and st.button("⚡ Распознать через ИИ"):
                with st.spinner("Извлечение данных через Qwen3-VL-Plus..."):
                    extracted_data = ai_service.extract_batch_info_from_image(sticker_file.getvalue())

        col_b1, col_b2, col_b3 = st.columns(3)
        report_no = col_b1.text_input("Report No.", value="QC-" + st.session_state.auth_user.id[:5].upper())
        client_name = col_b2.text_input("Клиент / Заказчик")
        supplier_name = col_b3.text_input("Фабрика / Поставщик", value=extracted_data.get("supplier", ""))
        
        col_b4, col_b5, col_b6 = st.columns(3)
        order_no = col_b4.text_input("PO / Order No.", value=extracted_data.get("order_no", ""))
        item_name = col_b5.text_input("Наименование товара", value=extracted_data.get("item_name", ""))
        sku = col_b6.text_input("SKU / Артикул", value=extracted_data.get("sku", ""))

    # 2. Размерная сетка (до 2 фото эталона)
    with st.container(border=True):
        st.markdown("**2. Размерная сетка (Garment Measurement Chart)**")
        use_spec = st.checkbox("Включить модуль размерной сетки для этой инспекции", value=True)
        
        if use_spec:
            if not st.session_state.reference_spec:
                st.caption("Загрузите 1 или 2 фото листа спецификации из техпака.")
                spec_imgs = st.file_uploader(
                    "Фото эталонной спецификации (до 2 фото)", 
                    type=["jpg", "png", "jpeg"], 
                    accept_multiple_files=True, 
                    key="spec_upload"
                )
                
                if spec_imgs:
                    if len(spec_imgs) > 2:
                        st.warning("Загружено больше 2 фото. Первые 2 будут взяты в обработку.")
                        spec_imgs = spec_imgs[:2]
                    
                    cols = st.columns(len(spec_imgs))
                    for idx, img_file in enumerate(spec_imgs):
                        with cols[idx]:
                            st.image(img_file, caption=f"Фото эталона #{idx+1}", use_container_width=True)

                    if st.button("⚡ Оцифровать эталон (объединить фото)", type="primary", use_container_width=True):
                        with st.spinner("Нейросеть считывает и объединяет данные с фото..."):
                            bytes_list = [f.getvalue() for f in spec_imgs]
                            spec = ai_service.digitize_spec_sheet(bytes_list)
                            if spec and "poms" in spec:
                                st.session_state.reference_spec = spec
                                st.rerun()
                            else:
                                st.error("Не удалось извлечь таблицу. Проверьте четкость фото.")
            else:
                spec = st.session_state.reference_spec
                st.success(f"Эталон загружен. Размеры: {', '.join(spec.get('sizes', []))} ({spec.get('unit', 'cm')})")
                if st.button("Сбросить эталонную сетку"):
                    st.session_state.reference_spec = None
                    st.session_state.measured_samples = []
                    st.rerun()

                s_num = len(st.session_state.measured_samples) + 1
                st.markdown(f"##### ➕ Замер образца #{s_num}")
                selected_size = st.selectbox("Размер проверяемого образца", spec.get("sizes", []), key=f"sz_{s_num}")
                
                inputs = {}
                with st.form(key=f"sample_form_{s_num}"):
                    for pom in spec.get("poms", []):
                        p_code = pom.get("code", "")
                        p_desc = pom.get("description", "")
                        ideal_val = pom.get("specs", {}).get(selected_size, 0.0)
                        t_min = pom.get("tol_minus", 1.0)
                        t_plus = pom.get("tol_plus", 1.0)

                        c1, c2, c3 = st.columns([1, 4, 3])
                        c1.write(f"**{p_code}**")
                        c2.write(f"{p_desc} (±{t_plus})")
                        inputs[p_code] = {
                            "val": c3.text_input(f"Замер {p_code}", placeholder=str(ideal_val), label_visibility="collapsed"),
                            "ideal": ideal_val,
                            "desc": p_desc,
                            "t_min": t_min,
                            "t_plus": t_plus
                        }
                    
                    if st.form_submit_button("✓ Проверить образец"):
                        sample_items = []
                        for code, item_meta in inputs.items():
                            raw = item_meta["val"].strip().replace(",", ".")
                            act = float(raw) if raw else item_meta["ideal"]
                            diff = round(act - item_meta["ideal"], 2)
                            is_fail = (diff < -item_meta["t_min"]) or (diff > item_meta["t_plus"])
                            sample_items.append({
                                "code": code,
                                "desc": item_meta["desc"],
                                "actual": act,
                                "diff": diff,
                                "status": "FAIL" if is_fail else "PASS"
                            })
                        
                        st.session_state.measured_samples.append({
                            "sample_no": s_num,
                            "size": selected_size,
                            "items": sample_items
                        })
                        st.rerun()

                if st.session_state.measured_samples:
                    st.markdown("**Готовые замеры в отчете:**")
                    for smp in st.session_state.measured_samples:
                        fails = sum(1 for it in smp["items"] if it["status"] == "FAIL")
                        st.write(f"- Образец #{smp['sample_no']} ({smp['size']}): " + (f"⚠️ Отклонений: {fails}" if fails else "✅ В норме"))

    # 3. Фотоматериалы (до 12 фото на блок, чистая галерея превью)
    photos_data = {"products": [], "packing": [], "labels": [], "defects": []}
    with st.container(border=True):
        st.markdown("**3. Фотоматериалы инспекции (до 12 фото на секцию)**")
        sec_names = {
            "products": "Продукция (Products)",
            "packing": "Упаковка и маркировка (Packaging)",
            "labels": "Стикеры и штрихкоды (Labels)",
            "defects": "Дефекты (Defects)"
        }
        for sec_key, sec_title in sec_names.items():
            st.markdown(f"##### {sec_title}")
            files = st.file_uploader(
                f"Загрузить фото для секции '{sec_title}'", 
                type=["jpg", "png", "jpeg"], 
                accept_multiple_files=True, 
                key=f"f_{sec_key}",
                label_visibility="collapsed"
            )
            
            if files:
                if len(files) > 12:
                    st.warning(f"Выбрано {len(files)} фото. В отчет попадут первые 12.")
                    files = files[:12]
                
                preview_cols = st.columns(4)
                for idx, fl in enumerate(files):
                    with preview_cols[idx % 4]:
                        st.image(fl, use_container_width=True)
                    
                    photos_data[sec_key].append({"bytes": fl.getvalue()})
            st.divider()

    # 4. Комментарии инспектора
    with st.container(border=True):
        st.markdown("**4. Комментарии**")
        raw_comments = st.text_area(
            "Комментарии", 
            value=st.session_state.comments_val,
            placeholder="Рабочие заметки инспектора...",
            label_visibility="collapsed"
        )
        if st.button("✨ Рецензировать с ИИ", key="btn_ai_comments"):
            if raw_comments.strip():
                with st.spinner("Рецензирование..."):
                    st.session_state.comments_val = ai_service.polish_qc_text(raw_comments, field_type="comments")
                    st.rerun()

    # 5. Общий итог проверки
    with st.container(border=True):
        st.markdown("**5. Общий итог проверки**")
        default_conclusion = st.session_state.conclusion_val if st.session_state.conclusion_val else st.session_state.comments_val
        final_comments = st.text_area(
            "Итоговые выводы для отчета", 
            value=default_conclusion,
            placeholder="Итоговое заключение для отчета...",
            label_visibility="collapsed"
        )
        if st.button("✨ Рецензировать с ИИ", key="btn_ai_conclusion"):
            if final_comments.strip():
                with st.spinner("Рецензирование..."):
                    st.session_state.conclusion_val = ai_service.polish_qc_text(final_comments, field_type="conclusion")
                    st.rerun()

        result_status = st.selectbox("Финальный результат инспекции", ["PASSED", "PENDING", "REJECTED"])

    # Сборка Excel и сохранение
    if st.button("🚀 Сформировать и сохранить отчет Excel", type="primary", use_container_width=True):
        rep_payload = {
            "report_no": report_no,
            "company_name": st.secrets.get("DEFAULT_COMPANY_NAME", "BIZNESKIT"),
            "client_name": client_name,
            "supplier_name": supplier_name,
            "order_no": order_no,
            "item_name": item_name,
            "sku": sku,
            "inspector_id": st.session_state.auth_user.email,
            "result_status": result_status,
            "comments": final_comments
        }
        
        try:
            with st.spinner("Генерация официального Excel-акта в высоком разрешении..."):
                excel_bytes = create_full_qc_report(
                    report_data=rep_payload,
                    photos_by_section=photos_data,
                    spec_data=st.session_state.reference_spec,
                    measured_samples=st.session_state.measured_samples
                )
                st.session_state.last_excel_bytes = excel_bytes
                st.session_state.last_excel_name = _safe_download_name(report_no, result_status)

                cloud_saved = save_report_with_excel(
                    {
                        "inspector_id": st.session_state.auth_user.id,
                        "report_no": report_no,
                        "client_name": client_name,
                        "supplier_name": supplier_name,
                        "order_no": order_no,
                        "item_name": item_name,
                        "sku": sku,
                        "result_status": result_status,
                        "report_data": rep_payload,
                    },
                    excel_bytes,
                )

                if cloud_saved:
                    st.success("Отчет успешно сформирован и сохранён в архив.")
                else:
                    st.success("Отчет успешно сформирован.")
                    st.warning(
                        "Облачное сохранение не удалось. Файл Excel можно скачать в этой сессии. "
                        "В архиве он может отсутствовать, пока сохранение не пройдёт успешно."
                    )
        except Exception:
            st.error("Не удалось сформировать Excel-файл. Проверьте фото и данные отчёта.")

    if st.session_state.last_excel_bytes:
        st.download_button(
            label="📥 Скачать QC Report (.xlsx)",
            data=st.session_state.last_excel_bytes,
            file_name=st.session_state.last_excel_name,
            mime=EXCEL_MIME,
            use_container_width=True,
            key="dl_last_excel",
        )

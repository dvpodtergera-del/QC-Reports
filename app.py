from __future__ import annotations

import sys
import os
import io
import urllib.request
import uuid
from datetime import datetime, date
import base64
import streamlit as st
import streamlit.components.v1 as components
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
# РАСЧЁТ ЗАМЕРОВ
# ==============================================================================
MEASURE_STATUS_LABELS = {
    "PASS": "в допуске",
    "FAIL": "вне допуска",
    "NOT_MEASURED": "не измерено",
    "INVALID": "ошибка ввода",
    "NO_TARGET": "нет эталона/допуска",
}


def _to_number(value):
    """Число из значения спецификации или None, если это не число."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError):
        return None


def evaluate_measurement(raw, ideal, tol_minus, tol_plus) -> dict:
    """Оценивает один замер. Пустое поле никогда не считается PASS."""
    text = "" if raw is None else str(raw).strip()
    if not text:
        return {"status": "NOT_MEASURED", "actual": None, "diff": None}
    try:
        actual = float(text.replace(",", "."))
    except ValueError:
        return {"status": "INVALID", "actual": None, "diff": None, "raw": text}

    ideal_n = _to_number(ideal)
    t_min = _to_number(tol_minus)
    t_plus = _to_number(tol_plus)
    if ideal_n is None or t_min is None or t_plus is None:
        return {"status": "NO_TARGET", "actual": actual, "diff": None}

    diff = round(actual - ideal_n, 2)
    # Граница допуска включительно; небольшой запас против ошибок округления float.
    eps = 1e-9
    is_fail = diff < -abs(t_min) - eps or diff > abs(t_plus) + eps
    return {"status": "FAIL" if is_fail else "PASS", "actual": actual, "diff": diff}


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
# МОДУЛЬ ГЕНЕРАЦИИ EXCEL (HD-ФОТОСЕТКА С ПОДПИСЯМИ СЕКЦИЙ)
# ==============================================================================

def _size_sort_key(value):
    import re
    text = str(value).strip().upper().replace(' ', '')
    order = {'XXXS':0, '3XS':0, 'XXS':1, '2XS':1, 'XS':2, 'S':3, 'M':4,
             'L':5, 'XL':6, 'XXL':7, '2XL':7, 'XXXL':8, '3XL':8, '4XL':9, '5XL':10, '6XL':11}
    if text in order:
        return (0, order[text], '')
    numbers = re.findall(r'\d+(?:[.,]\d+)?', text)
    if numbers:
        return (1, tuple(float(n.replace(',', '.')) for n in numbers), text)
    return (2, (), text)


def _frame_report(ws, last_row, right=10):
    from copy import copy
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins
    edge = Side(style='medium', color='334155')
    for row in ws.iter_rows(min_row=2, max_row=last_row, min_col=2, max_col=right):
        for cell in row:
            if cell.fill.patternType is None:
                cell.fill = PatternFill('solid', fgColor='FFFFFF')
            b = copy(cell.border)
            if cell.column == 2: b.left = edge
            if cell.column == right: b.right = edge
            if cell.row == 2: b.top = edge
            if cell.row == last_row: b.bottom = edge
            cell.border = b
    ws.sheet_view.showGridLines = False
    ws.print_area = f'B2:{get_column_letter(right)}{last_row}'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(left=.25,right=.25,top=.35,bottom=.35,header=.15,footer=.15)


def _photo_section(ws, row, title, photos):
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
    from openpyxl.utils.units import pixels_to_EMU
    from openpyxl.utils import get_column_letter
    ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=10)
    c=ws.cell(row,2,title)
    c.fill=PatternFill('solid',fgColor='334155')
    c.font=Font(name='Calibri',size=10,bold=True,color='FFFFFF')
    c.alignment=Alignment(vertical='center',indent=1)
    ws.row_dimensions[row].height=24
    row+=1
    gap=8
    slots=((2,4),(5,7),(8,10))
    widths=[sum(int((ws.column_dimensions[get_column_letter(c)].width or 15)*7+.5)+5 for c in range(left,right+1)) for left,right in slots]
    height=round((min(widths)-gap)*9/16)
    batches=(len(photos)+2)//3
    for batch,base in enumerate(range(0,len(photos),3)):
        bottom_gap=gap if batch<batches-1 else 0
        ws.row_dimensions[row].height=(height+bottom_gap)*.75
        for slot,(left,right) in enumerate(slots):
            ws.merge_cells(start_row=row,start_column=left,end_row=row,end_column=right)
            idx=base+slot
            if idx>=len(photos):continue
            left_inset=0 if slot==0 else gap//2
            right_inset=0 if slot==2 else gap//2
            width=widths[slot]-left_inset-right_inset
            with Image.open(io.BytesIO(photos[idx]['bytes'])) as source:
                image=ImageOps.exif_transpose(source).convert('RGB')
                image=ImageOps.fit(image,(width*4,height*4),method=Image.Resampling.LANCZOS,centering=(.5,.5))
                buffer=io.BytesIO();image.save(buffer,format='JPEG',quality=94)
            buffer.seek(0)
            picture=OpenPyXLImage(buffer)
            picture.anchor=TwoCellAnchor(editAs='oneCell',
                _from=AnchorMarker(col=left-1,row=row-1,colOff=pixels_to_EMU(left_inset),rowOff=0),
                to=AnchorMarker(col=right,row=row,colOff=-pixels_to_EMU(right_inset),rowOff=-pixels_to_EMU(bottom_gap)))
            ws.add_image(picture)
        row+=1
    return row



def _create_measurements_sheet(wb,spec,samples,report_data):
    from openpyxl.utils import get_column_letter
    ws=wb.create_sheet('Размерная сетка')
    poms=spec.get('poms') or []
    sizes=sorted({str(x) for x in spec.get('sizes',[])},key=_size_sort_key)
    if not sizes:sizes=sorted({str(k) for p in poms for k in (p.get('specs') or {})},key=_size_sort_key)
    groups={size:[x for x in samples if str(x.get('size',''))==size] for size in sorted({str(x.get('size','')) for x in samples},key=_size_sort_key)}
    right=max(8,5+len(sizes),max((6+3*len(group) for group in groups.values()),default=0))
    if right>16384:raise ValueError('Слишком много образцов для одного листа Excel.')
    ws.column_dimensions['A'].width=3
    ws.column_dimensions['B'].width=11
    ws.column_dimensions['C'].width=42
    for col in range(4,right+1):ws.column_dimensions[get_column_letter(col)].width=15
    thin=Side(style='thin',color='CBD5E1')
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    colors={'PASS':('DCFCE7','166534'),'FAIL':('FEE2E2','991B1B')}
    def put(r,c,value,header=False,status=None):
        x=ws.cell(r,c,value);x.border=border
        bg,fg=('475569','FFFFFF') if header else colors.get(status,('FEF3C7','92400E') if status else ('F8FAFC' if r%2 else 'FFFFFF','0F172A'))
        x.fill=PatternFill('solid',fgColor=bg)
        x.font=Font(name='Calibri',size=13,bold=header or status is not None,color=fg)
        x.alignment=Alignment(horizontal='left' if c==3 else 'center',vertical='center',textRotation=0,wrap_text=not header)
        if isinstance(value,(int,float)):x.number_format='0.##'
        return x
    def band(r,text,end=right,color='334155'):
        ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=end)
        x=ws.cell(r,2,text);x.font=Font(name='Calibri',size=14.3,bold=True,color='FFFFFF')
        x.fill=PatternFill('solid',fgColor=color)
        x.alignment=Alignment(vertical='center',indent=1,textRotation=0,wrap_text=True)
        ws.row_dimensions[r].height=38
    band(2,f"РАЗМЕРНАЯ СЕТКА — {report_data.get('report_no','')} — {spec.get('unit','cm')}",color='1E293B')
    band(4,'ЭТАЛОН — оцифрованная спецификация',end=max(8,5+len(sizes)))
    for c,value in enumerate(['Код','Описание','Допуск −','Допуск +']+sizes,2):put(5,c,value,True)
    ws.row_dimensions[5].height=36
    row=6
    for pom in poms:
        refs={str(k):v for k,v in (pom.get('specs') or {}).items()}
        for c,value in enumerate([pom.get('code',''),pom.get('description',''),_to_number(pom.get('tol_minus')),_to_number(pom.get('tol_plus'))]+[_to_number(refs.get(size)) for size in sizes],2):put(row,c,value)
        ws.row_dimensions[row].height=38;row+=1
    row+=2;band(row,'ПРОВЕРЕННЫЕ ОБРАЗЦЫ');row+=2
    if not groups:band(row,'Проверенные образцы пока отсутствуют',color='64748B');row+=1
    for size,group in groups.items():
        end=6+3*len(group)
        band(row,f'РАЗМЕР {size} — образцов проверено: {len(group)}',end=end);row+=1
        for c,label in enumerate(['Код','Описание','Критических','Допуск −','Допуск +'],2):
            put(row,c,label,True);put(row+1,c,'',True)
        for idx,sample in enumerate(group):
            col=7+idx*3
            ws.merge_cells(start_row=row,start_column=col,end_row=row,end_column=col+2)
            put(row,col,f"Образец #{sample.get('sample_no','')}",True)
            for off,label in enumerate(['Эталон','Факт','Δ']):put(row+1,col+off,label,True)
        ws.row_dimensions[row].height=36;ws.row_dimensions[row+1].height=34;row+=2
        pom_map={str(p.get('code','')):p for p in poms};codes=list(pom_map)
        maps=[{str(item.get('code','')):item for item in sample.get('items',[])} for sample in group]
        for items in maps:
            for code in items:
                if code not in codes:codes.append(code)
        for code in codes:
            pom=pom_map.get(code,{})
            desc=pom.get('description') or next((items[code].get('desc','') for items in maps if code in items),'')
            critical=sum(items.get(code,{}).get('status')=='FAIL' for items in maps)
            refs={str(k):v for k,v in (pom.get('specs') or {}).items()}
            for c,value in enumerate([code,desc,critical,_to_number(pom.get('tol_minus')),_to_number(pom.get('tol_plus'))],2):put(row,c,value,status='FAIL' if c==4 and critical else None)
            for idx,items in enumerate(maps):
                item=items.get(code,{});status=item.get('status','NOT_MEASURED');col=7+idx*3
                for off,value in enumerate([_to_number(refs.get(size)),item.get('actual') if item.get('actual') is not None else '—',item.get('diff') if item.get('diff') is not None else '—']):put(row,col+off,value,status=status if off else None)
            ws.row_dimensions[row].height=max(38,20*((len(str(desc))//38)+1));row+=1
        good,bad,incomplete,details=[],[],[],[]
        for sample,items in zip(group,maps):
            num=sample.get('sample_no','?')
            failed=[code for code in codes if items.get(code,{}).get('status')=='FAIL']
            if failed:
                bad.append(num)
                for code in failed:
                    item=items[code];pom=pom_map.get(code,{})
                    refs={str(k):v for k,v in (pom.get('specs') or {}).items()}
                    delta=item.get('diff');delta_text=f'{delta:+g}' if isinstance(delta,(float,int)) else str(delta)
                    description=pom.get('description') or item.get('desc','')
                    details.append(f"#{num} · {code} {description}: факт {item.get('actual','—')}; эталон {refs.get(size,'—')}; Δ {delta_text}; допуск −{pom.get('tol_minus','?')} / +{pom.get('tol_plus','?')}")
            elif codes and all(items.get(code,{}).get('status')=='PASS' for code in codes):good.append(num)
            else:incomplete.append(num)
        row+=1;band(row,f'СТАТИСТИКА ПО РАЗМЕРУ {size}',end=end,color='475569');row+=1
        summary=[f'Проверено: {len(group)} · В допуске: {len(good)} · Вне допуска: {len(bad)}',
                 'Образцы с расхождениями: '+(', '.join('#'+str(num) for num in bad) or 'нет')]
        if incomplete:summary.append('Не завершена оценка: '+', '.join('#'+str(num) for num in incomplete))
        for text in summary:
            ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=end)
            x=ws.cell(row,2,text);x.font=Font(name='Calibri',size=13,color='334155')
            x.alignment=Alignment(vertical='center',wrap_text=True,textRotation=0,indent=1)
            ws.row_dimensions[row].height=max(34,22*(len(text)//max(70,end*9)+1));row+=1
        row+=2
    ws.freeze_panes='G6'
    ws.sheet_view.showGridLines=False
    ws.sheet_view.zoomScale=100
    ws.print_area=f'B2:{get_column_letter(right)}{row-1}'
    ws.sheet_properties.pageSetUpPr.fitToPage=False
    ws.page_setup.orientation='landscape';ws.page_setup.scale=100
    ws.page_setup.fitToWidth=0;ws.page_setup.fitToHeight=0
    return ws


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

    # Measurements are stored on their own sheet, not across the main report.
    if spec_data:
        _create_measurements_sheet(wb, spec_data, measured_samples or [], report_data)

    # 4. Фотоматериалы: три фотографии в ряд, сгруппированные по секциям
    section_titles = {
        "products": "PRODUCT WORKMANSHIP PHOTOS",
        "packing": "PACKAGING & CARTON DETAILS",
        "labels": "LABELS & BARCODES VERIFICATION",
        "defects": "DEFECT DETAILS & LOG"
    }

    for letter in 'BCDEFGHIJ':
        ws.column_dimensions[letter].width = 15
    for key, title in section_titles.items():
        if photos_by_section.get(key):
            curr_row = _photo_section(ws, curr_row, title, photos_by_section[key])

    # 5. Отдельные блоки комментариев инспектора и заключения
    def add_text_block(title: str, text: str, row: int) -> int:
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        ws.row_dimensions[row].height = 22
        header_cell = ws.cell(row=row, column=2, value=title)
        header_cell.fill = PatternFill(start_color=c_header, end_color=c_header, fill_type="solid")
        header_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        header_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        apply_range_style(ws, 2, row, 9, row, border=thin_border)
        row += 1
        lines = str(text or "—").splitlines() or ["—"]
        for line in lines:
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
            cell = ws.cell(row=row, column=2, value=line)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.font = Font(name="Calibri", size=9, color="0F172A")
            apply_range_style(ws, 2, row, 9, row, border=thin_border)
            ws.row_dimensions[row].height = 20
            row += 1
        return row + 1

    curr_row = add_text_block("INSPECTOR COMMENTS", report_data.get("inspector_comments", report_data.get("comments", "")), curr_row)
    curr_row = add_text_block("INSPECTION CONCLUSION", report_data.get("conclusion", ""), curr_row)

    # Итог проверки дополнительно дублируется внизу отчёта.
    ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=9)
    ws.row_dimensions[curr_row].height = 28
    result_cell = ws.cell(row=curr_row, column=2, value=f"FINAL INSPECTION RESULT: {res_status}")
    result_cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    result_cell.alignment = Alignment(horizontal="center", vertical="center")
    if res_status == "PASSED":
        result_cell.fill = PatternFill(start_color=c_pass_fg, end_color=c_pass_fg, fill_type="solid")
    elif res_status == "REJECTED":
        result_cell.fill = PatternFill(start_color=c_fail_fg, end_color=c_fail_fg, fill_type="solid")
    else:
        result_cell.fill = PatternFill(start_color=c_pend_fg, end_color=c_pend_fg, fill_type="solid")
    apply_range_style(ws, 2, curr_row, 9, curr_row, border=thin_border)

    # Extend full-width headers and text blocks to the photo-grid edge.
    for merged in list(ws.merged_cells.ranges):
        if merged.max_col == 9:
            left, top, bottom = merged.min_col, merged.min_row, merged.max_row
            ws.unmerge_cells(str(merged))
            ws.merge_cells(start_row=top, end_row=bottom, start_column=left, end_column=10)
    _frame_report(ws, curr_row)
    out = io.BytesIO()
    for sheet in wb.worksheets:
        for cells in sheet:
            for cell in cells:
                if cell.comment is not None: cell.comment = None
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


def append_photos_to_existing_excel(excel_bytes: bytes, new_photos: list[tuple[str, bytes]], report_data: dict) -> bytes:
    """Пересобирает фотоблоки в исходных секциях: сохраняет замеры, распределение фото и заключение."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active

    updates = {
        "C4": report_data.get("report_no", ""),
        "G4": report_data.get("inspection_date", ""),
        "C5": report_data.get("client_name", ""),
        "G5": report_data.get("inspector_id", ""),
        "C6": report_data.get("supplier_name", ""),
        "G6": report_data.get("order_no", ""),
        "C7": f"{report_data.get('item_name', '')} / {report_data.get('sku', '')}",
        "G7": report_data.get("result_status", ""),
    }
    for cell_ref, value in updates.items():
        try:
            ws[cell_ref] = str(value or "")
        except Exception:
            pass

    section_order = [
        ("PRODUCT WORKMANSHIP PHOTOS", "Продукция"),
        ("PACKAGING & CARTON DETAILS", "Упаковка и маркировка"),
        ("LABELS & BARCODES VERIFICATION", "Стикеры и штрихкоды"),
        ("DEFECT DETAILS & LOG", "Дефекты"),
    ]
    title_to_key = {}
    for key, russian_title in section_order:
        title_to_key[key] = key
        title_to_key[russian_title] = key
    grouped = {key: [] for key, _ in section_order}

    # Восстанавливаем все существующие фото из Excel, сохраняя их секцию.
    for item in extract_images_from_excel(excel_bytes):
        section_title = item.get("section", "")
        key = title_to_key.get(section_title)
        if key:
            grouped[key].append({"bytes": item["bytes"], "caption": item.get("label", "")})
        else:
            # Если старый файл был создан прежней версией и секция не распознана,
            # не теряем фотографии: помещаем их в секцию продукции.
            grouped["PRODUCT WORKMANSHIP PHOTOS"].append({"bytes": item["bytes"], "caption": item.get("label", "")})

    # Новые фотографии закрепляются за выбранным разделом.
    for section_title, raw in new_photos:
        key = title_to_key.get(section_title)
        if key and raw:
            grouped[key].append({"bytes": raw, "caption": ""})

    # Находим начало фотоблока/заключения. Всё выше этой строки (включая замеры)
    # сохраняем. Фотографии и заключение ниже полностью формируем заново.
    start_candidates = []
    for row in range(1, ws.max_row + 1):
        val = str(ws.cell(row=row, column=2).value or "").strip()
        if val in PHOTO_SECTION_TITLES or val in (
            "INSPECTION COMMENTS & CONCLUSION",
            "UPDATED COMMENTS AND CONCLUSION",
            "ADDITIONAL PHOTOS — UPDATED REPORT",
        ):
            start_candidates.append(row)
    start_row = min(start_candidates) if start_candidates else ws.max_row + 1

    # Удаляем изображения (они будут вставлены заново в корректную сетку 3 в ряд).
    ws._images = []
    if start_row <= ws.max_row:
        for merged in list(ws.merged_cells.ranges):
            if merged.min_row >= start_row:
                ws.unmerge_cells(str(merged))
        ws.delete_rows(start_row, ws.max_row - start_row + 1)

    c_header, c_sub = '334155', '475569'
    thin_border = Border(*([Side(style='thin',color='CBD5E1')]*4))
    for letter in 'BCDEFGHIJ':
        ws.column_dimensions[letter].width=15
    row=start_row
    for key,title in section_order:
        if grouped[key]:
            row=_photo_section(ws,row,key,grouped[key])

    # Обновляем отдельные текстовые блоки и дублируем результат внизу.
    row += 1
    def add_footer_text(title: str, text: str, row_num: int) -> int:
        ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=10)
        header = ws.cell(row=row_num, column=2, value=title)
        header.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        header.fill = PatternFill(start_color=c_header, end_color=c_header, fill_type="solid")
        header.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row_num].height = 22
        apply_range_style(ws, 2, row_num, 10, row_num, border=thin_border)
        row_num += 1
        for line in (str(text or "—").splitlines() or ["—"]):
            ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=10)
            cell = ws.cell(row=row_num, column=2, value=line)
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.font = Font(name="Calibri", size=9, color="0F172A")
            apply_range_style(ws, 2, row_num, 10, row_num, border=thin_border)
            ws.row_dimensions[row_num].height = 20
            row_num += 1
        return row_num + 1

    row = add_footer_text("INSPECTOR COMMENTS", report_data.get("inspector_comments", report_data.get("comments", "")), row)
    row = add_footer_text("INSPECTION CONCLUSION", report_data.get("conclusion", ""), row)
    result = str(report_data.get("result_status") or "PENDING").upper()
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=10)
    result_cell = ws.cell(row=row, column=2, value=f"FINAL INSPECTION RESULT: {result}")
    result_cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    result_cell.alignment = Alignment(horizontal="center", vertical="center")
    result_color = "166534" if result == "PASSED" else "991B1B" if result == "REJECTED" else "92400E"
    result_cell.fill = PatternFill(start_color=result_color, end_color=result_color, fill_type="solid")
    ws.row_dimensions[row].height = 28
    apply_range_style(ws, 2, row, 10, row, border=thin_border)

    _frame_report(ws, row)
    output = io.BytesIO()
    for sheet in wb.worksheets:
        for cells in sheet:
            for cell in cells:
                if cell.comment is not None: cell.comment = None
    wb.save(output)
    wb.close()
    return output.getvalue()


PHOTO_SECTION_TITLES = {
    "PRODUCT WORKMANSHIP PHOTOS": "Продукция",
    "PACKAGING & CARTON DETAILS": "Упаковка и маркировка",
    "LABELS & BARCODES VERIFICATION": "Стикеры и штрихкоды",
    "DEFECT DETAILS & LOG": "Дефекты",
}


def extract_images_from_excel(excel_bytes: bytes) -> list[dict]:
    """Извлекает фотографии Excel вместе с подписью секции, если она есть на листе."""
    images: list[dict] = []
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
        for worksheet in workbook.worksheets:
            section_rows: list[tuple[int, str]] = []
            for row in worksheet.iter_rows(min_col=2, max_col=2):
                for cell in row:
                    raw_title = str(cell.value or "").strip()
                    if raw_title in PHOTO_SECTION_TITLES:
                        section_rows.append((cell.row, PHOTO_SECTION_TITLES[raw_title]))

            pending: list[tuple[int, bytes]] = []
            for embedded_image in getattr(worksheet, "_images", []):
                try:
                    raw = embedded_image._data()
                    if not raw:
                        continue
                    with Image.open(io.BytesIO(raw)) as test_image:
                        test_image.verify()
                    anchor = getattr(embedded_image, "anchor", None)
                    from_cell = getattr(anchor, "_from", None)
                    img_row = int(getattr(from_cell, "row", 0) or 0) + 1
                    pending.append((img_row, raw))
                except Exception:
                    continue

            pending.sort(key=lambda item: item[0])
            counters: dict[str, int] = {}
            title_aliases = {
                "PRODUCT WORKMANSHIP PHOTOS": "Продукция",
                "PACKAGING & CARTON DETAILS": "Упаковка и маркировка",
                "LABELS & BARCODES VERIFICATION": "Стикеры и штрихкоды",
                "DEFECT DETAILS & LOG": "Дефекты",
            }
            for img_row, raw in pending:
                section = ""
                for header_row, header_title in section_rows:
                    if header_row <= img_row:
                        section = header_title
                # Совместимость со старыми Excel, где фото добавлялись в конец
                # с текстом секции в ячейке B перед фотографией.
                for rr in range(max(1, img_row - 3), img_row + 1):
                    cell_text = str(worksheet.cell(row=rr, column=2).value or "").strip()
                    for english_title, russian_title in title_aliases.items():
                        if cell_text.startswith(english_title):
                            section = russian_title
                            break
                    if section in title_aliases.values():
                        break
                counters[section] = counters.get(section, 0) + 1
                label = f"{section} — {counters[section]}" if section else f"Фото {counters[section]}"
                images.append({"bytes": raw, "label": label, "section": section})
        workbook.close()
    except Exception:
        return []
    return images


@st.dialog("Просмотр фотографии", width="large")
def show_photo_viewer():
    """Модальное окно для просмотра и перелистывания фотографий отчета."""
    photos = st.session_state.get("photo_viewer_images", [])
    if not photos:
        st.info("Фотографии не найдены.")
        return

    index = int(st.session_state.get("photo_viewer_index", 0)) % len(photos)
    st.session_state.photo_viewer_index = index

    # Диалог занимает около 80% ширины окна приложения; фото не растягивается по высоте.
    st.markdown(
        """<style>
        div[role="dialog"] { width: 98vw !important; max-width: 98vw !important; min-width: 0 !important; }
        div[role="dialog"] img { max-height: 82dvh; object-fit: contain; }
        </style>""",
        unsafe_allow_html=True,
    )
    st.caption(f"Фото {index + 1} из {len(photos)}")
    st.image(photos[index], use_container_width=True)

    if st.button("✕ Закрыть просмотр", key="photo_viewer_close", use_container_width=True):
        st.session_state.photo_viewer_open = False
        st.rerun()

    prev_col, counter_col, next_col = st.columns([1, 2, 1])
    with prev_col:
        if st.button("← Назад", key="photo_viewer_prev", use_container_width=True):
            st.session_state.photo_viewer_index = (index - 1) % len(photos)
            st.rerun()
    with counter_col:
        st.markdown(
            f"<div style='text-align:center;padding-top:8px'>{index + 1} / {len(photos)}</div>",
            unsafe_allow_html=True,
        )
    with next_col:
        if st.button("Вперёд →", key="photo_viewer_next", use_container_width=True):
            st.session_state.photo_viewer_index = (index + 1) % len(photos)
            st.rerun()


def render_clickable_photo_gallery(photos: list, height: int | None = None) -> None:
    """Кликабельная галерея с увеличением и переключением стрелками."""
    if not photos:
        st.caption("Фотографии не найдены.")
        return
    items = []
    for i, photo in enumerate(photos):
        if isinstance(photo, dict):
            raw = photo.get("bytes")
            label = str(photo.get("label") or f"Фото {i + 1}")
            section = str(photo.get("section") or "")
        else:
            raw = photo
            label = f"Фото {i + 1}"
            section = ""
        if not raw:
            continue
        try:
            with Image.open(io.BytesIO(raw)) as im:
                fmt = (im.format or "JPEG").upper()
            mime = "image/png" if fmt == "PNG" else "image/webp" if fmt == "WEBP" else "image/jpeg"
            items.append({
                "src": "data:" + mime + ";base64," + base64.b64encode(raw).decode("ascii"),
                "label": label,
                "section": section,
            })
        except Exception:
            continue
    if not items:
        st.caption("Не удалось подготовить фотографии для просмотра.")
        return
    import json
    import math
    n_sections = len({item["section"] for item in items if item["section"]})
    photo_rows = max(1, math.ceil(len(items) / 4))
    if height is None:
        height = min(560, 28 + n_sections * 26 + photo_rows * 138)
    gallery_json = json.dumps(items, ensure_ascii=False).replace("</", "<\\/")
    html = f'''<!doctype html><html><head><meta charset="utf-8"><style>
*{{box-sizing:border-box}} body{{margin:0;font-family:Arial,sans-serif;color:#f5f5f5;background:transparent}}
.gallery{{height:{height-8}px;overflow:auto;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:6px 8px;padding:2px;align-content:start}}
.sec{{grid-column:1/-1;font-size:13px;font-weight:700;color:#e5e7eb;padding:4px 2px 0;margin:0}}
.thumb{{border:0;background:transparent;color:inherit;cursor:pointer;text-align:center;padding:0;min-width:0}}
.thumb img{{width:100%;height:110px;object-fit:cover;border-radius:8px;display:block;background:#222}}
.thumb span{{display:block;font-size:12px;padding-top:3px;color:#b7b7b7}}
.overlay{{display:none;position:fixed;inset:0;width:100%;height:100%;background:rgba(8,8,12,.98);z-index:9999;padding:8px;flex-direction:column;align-items:center;justify-content:center}}
.overlay.open{{display:flex}} .top{{width:100%;display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:10px}}
.action{{border:1px solid #666;border-radius:8px;padding:9px 13px;background:#222;color:white;cursor:pointer;font-size:15px}}
.photo-wrap{{position:relative;display:flex;align-items:center;justify-content:center;width:100%;flex:1;min-height:0}}
.big{{max-width:100%;max-height:100%;object-fit:contain;border-radius:5px}} .arrow{{position:absolute;top:50%;transform:translateY(-50%);font-size:25px;min-width:42px;padding:10px 5px;z-index:2}} #prev{{left:4px}} #next{{right:4px}}
@media(max-width:600px){{.gallery{{grid-template-columns:repeat(2,minmax(0,1fr))}}.thumb img{{height:96px}}.big{{max-width:100%}}.arrow{{min-width:30px}}}}
</style></head><body><div class="gallery" id="gallery"></div>
<div class="overlay" id="overlay" role="dialog" aria-modal="true"><div class="top"><span id="counter">Фото 1 из {len(items)}</span><button class="action" id="close">✕ Закрыть</button></div>
<div class="photo-wrap"><button class="action arrow" id="prev">←</button><img class="big" id="big" alt="Увеличенная фотография"><button class="action arrow" id="next">→</button></div></div>
<script>
const photos = {gallery_json}; let current = 0;
let root=document,host=null,keyDocument=document;
try{{if(window.parent.document!==document){{
 const d=window.parent.document;host=d.createElement('div');host.style.cssText='position:fixed;inset:0;z-index:2147483647;display:none';
 root=host.attachShadow({{mode:'open'}});const style=document.createElement('style');style.textContent=document.querySelector('style').textContent;
 root.append(style,document.getElementById('overlay').cloneNode(true));d.body.append(host);keyDocument=d;
}}}}catch(e){{console.warn(e)}}
const gallery=document.getElementById('gallery'),overlay=root.getElementById('overlay'),big=root.getElementById('big'),counter=root.getElementById('counter');
function fit(){{if(!big.naturalWidth||!overlay.classList.contains('open'))return;const area=big.parentElement,k=Math.min(area.clientWidth/big.naturalWidth,area.clientHeight/big.naturalHeight);big.style.width=`${{big.naturalWidth*k}}px`;big.style.height=`${{big.naturalHeight*k}}px`}}
big.addEventListener('load',fit);const ro=new ResizeObserver(fit);ro.observe(big.parentElement);
function closeViewer(){{overlay.classList.remove('open');if(host)host.style.display='none';if(document.fullscreenElement)document.exitFullscreen().catch(()=>{{}})}}

function show(i){{current=(i+photos.length)%photos.length;big.src=photos[current].src;counter.textContent=`Фото ${{current+1}} из ${{photos.length}}`;overlay.classList.add('open');if(host)host.style.display='block';else if(!document.fullscreenElement&&document.documentElement.requestFullscreen)document.documentElement.requestFullscreen().catch(()=>{{}});fit()}}
let lastSection = null;
photos.forEach((p,i)=>{{
  if(p.section && p.section !== lastSection){{
    const h=document.createElement('div'); h.className='sec'; h.textContent=p.section; gallery.appendChild(h); lastSection=p.section;
  }}
  const b=document.createElement('button');b.className='thumb';b.type='button';
  const img=document.createElement('img');img.alt=p.label;img.src=p.src;
  const label=document.createElement('span');label.textContent=p.label;
  b.appendChild(img);b.appendChild(label);b.addEventListener('click',()=>show(i));gallery.appendChild(b);
}});
root.getElementById('close').addEventListener('click',closeViewer);
root.getElementById('prev').addEventListener('click',()=>show(current-1));root.getElementById('next').addEventListener('click',()=>show(current+1));
function onKey(e){{if(!overlay.classList.contains('open'))return;if(e.key==='Escape')closeViewer();if(e.key==='ArrowLeft')show(current-1);if(e.key==='ArrowRight')show(current+1)}}
keyDocument.addEventListener('keydown',onKey);
overlay.addEventListener('click',e=>{{if(e.target!==big&&!e.target.closest('button'))closeViewer()}});
function cleanup(){{ro.disconnect();keyDocument.removeEventListener('keydown',onKey);if(host)host.remove()}}
window.addEventListener('pagehide',cleanup);
if(host&&window.frameElement){{const frame=window.frameElement;const observer=new window.parent.MutationObserver(()=>{{if(!frame.isConnected){{cleanup();observer.disconnect()}}}});observer.observe(keyDocument.body,{{childList:true,subtree:true}});window.addEventListener('pagehide',()=>observer.disconnect())}}
</script></body></html>'''
    components.html(html, height=max(300, min(height, 560)), scrolling=False)


def _make_report_number() -> str:
    """Читаемый уникальный номер отчёта: QC-YYYYMMDD-HHMMSS-XXXX."""
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    suffix = uuid.uuid4().hex[:4].upper()
    return f"QC-{stamp}-{suffix}"


def _start_new_report() -> None:
    """Создаёт чистый черновик, не затрагивая уже сохранённые отчёты."""
    auth_user = st.session_state.get("auth_user")
    # Очищаем состояние текущего черновика и ключи виджетов, сохраняя авторизацию.
    for key in list(st.session_state.keys()):
        if key != "auth_user":
            try:
                del st.session_state[key]
            except Exception:
                pass
    st.session_state.auth_user = auth_user
    st.session_state.reference_spec = None
    st.session_state.measured_samples = []
    st.session_state.comments_val = ""
    st.session_state.conclusion_val = ""
    st.session_state.last_excel_bytes = None
    st.session_state.last_excel_name = "qc_report.xlsx"
    st.session_state.draft_report_no = _make_report_number()
    st.session_state.selected_report_id = None


if "auth_user" not in st.session_state:
    st.session_state.auth_user = None
if "draft_report_no" not in st.session_state:
    st.session_state.draft_report_no = _make_report_number()
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
        auth_mode = st.radio("Режим входа", ["Вход", "Регистрация", "Восстановить пароль", "Повторить письмо подтверждения"], horizontal=False)
        email = st.text_input("Email", key="auth_email").strip().lower()
        pwd = st.text_input("Пароль", type="password", key="auth_password")

        if auth_mode == "Вход" and st.button("Войти", type="primary", use_container_width=True):
            if not email or not pwd:
                st.warning("Введите email и пароль.")
            else:
                try:
                    res = supabase.auth.sign_in_with_password({"email": email, "password": pwd})
                    if getattr(res, "user", None):
                        st.session_state.auth_user = res.user
                        st.rerun()
                    else:
                        st.error("Supabase не вернул пользователя. Проверьте подтверждение email в настройках проекта.")
                except Exception as e:
                    err = str(e)
                    if "email_not_confirmed" in err.lower() or "email not confirmed" in err.lower():
                        st.error("Email ещё не подтверждён. Выберите «Повторить письмо подтверждения». Если письмо не приходит, администратору нужно проверить SMTP и настройки подтверждения email в Supabase.")
                    elif "invalid login credentials" in err.lower():
                        st.error("Supabase отклонил email или пароль. Если вы только что зарегистрировались, возможно, email не подтверждён или регистрация не завершилась. Можно запросить письмо подтверждения ниже.")
                    else:
                        st.error(f"Не удалось войти: {err}")

        elif auth_mode == "Регистрация" and st.button("Создать аккаунт", type="primary", use_container_width=True):
            if not email or not pwd:
                st.warning("Введите email и пароль.")
            elif len(pwd) < 6:
                st.warning("Пароль должен содержать не менее 6 символов.")
            else:
                try:
                    res = supabase.auth.sign_up({"email": email, "password": pwd})
                    if getattr(res, "session", None) and getattr(res, "user", None):
                        st.session_state.auth_user = res.user
                        st.success("Аккаунт создан, вы вошли в систему.")
                        st.rerun()
                    elif getattr(res, "user", None):
                        st.success("Аккаунт создан, но Supabase не выдал сессию. Скорее всего, требуется подтверждение email. Проверьте папки «Спам» и «Промоакции». Если письмо не приходит, выберите «Повторить письмо подтверждения»; если и это не поможет, нужно настроить SMTP в Supabase.")
                    else:
                        st.error("Supabase не вернул данные нового пользователя. Проверьте настройки Auth в Supabase.")
                except Exception as e:
                    st.error(f"Не удалось создать аккаунт: {str(e)}")

        elif auth_mode == "Восстановить пароль" and st.button("Отправить ссылку для сброса пароля", type="primary", use_container_width=True):
            if not email:
                st.warning("Введите email, на который зарегистрирован аккаунт.")
            else:
                try:
                    supabase.auth.reset_password_for_email(email)
                    st.success("Запрос на сброс пароля отправлен. Проверьте входящие письма и папку «Спам». Если письма нет, проверьте SMTP в Supabase.")
                except Exception as e:
                    st.error(f"Не удалось отправить запрос: {str(e)}")

        elif auth_mode == "Повторить письмо подтверждения" and st.button("Отправить письмо ещё раз", type="primary", use_container_width=True):
            if not email:
                st.warning("Введите email, указанный при регистрации.")
            else:
                try:
                    supabase.auth.resend({"type": "signup", "email": email})
                    st.success("Запрос на повторную отправку выполнен. Проверьте входящие письма и папку «Спам». Если письмо не приходит, вероятно, нужно настроить SMTP или проверить настройки подтверждения email в Supabase.")
                except Exception as e:
                    st.error(f"Не удалось отправить письмо подтверждения: {str(e)}")
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

        selected_id = st.session_state.get("selected_report_id")
        selected_report = next(
            (r for r in report_list if str(r.get("id") or "") == str(selected_id)),
            None,
        )

        # Когда открыт отчёт, скрываем список архива: пользователь видит только сам отчёт.
        if report_list and selected_report is None:
            archive_action_col, _ = st.columns([2, 5])
            with archive_action_col:
                st.caption("Выберите отчёт, чтобы открыть его отдельно от списка.")
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
                                file_bytes = download_excel_from_storage(storage_path)

                            if file_bytes:
                                st.session_state[f"arch_excel_{rid}"] = file_bytes
                                st.session_state[f"arch_name_{rid}"] = _safe_download_name(
                                    str(rep.get("report_no") or "report"),
                                    str(rep.get("result_status") or ""),
                                )
                            else:
                                st.warning("Не удалось получить Excel из облачного хранилища.")

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

        # Просмотр выбранного отчёта вместо списка архива — не нужно прокручивать вниз.
        if selected_report:
            st.subheader(f"Отчёт: {selected_report.get('report_no') or '-'}")

            if st.button("← Назад в архив", key="close_report_view"):
                st.session_state.selected_report_id = None
                st.session_state.photo_viewer_open = False
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
                        "Комментарии инспектора",
                        value=str(report_data.get("inspector_comments") or report_data.get("comments") or ""),
                        height=120,
                    )
                    edited_conclusion = st.text_area(
                        "Заключение инспекции",
                        value=str(report_data.get("conclusion") or ""),
                        height=120,
                    )

                    st.markdown("#### Добавить фотографии в нужные разделы")
                    st.caption("Загружай фотографии отдельно в нужный блок. Существующие фото и замеры сохранятся; новые фото появятся в выбранных разделах и в Excel.")
                    edit_photo_uploads_by_section = {
                        "PRODUCT WORKMANSHIP PHOTOS": st.file_uploader(
                            "Продукция", type=["jpg", "jpeg", "png", "webp"],
                            accept_multiple_files=True, key=f"edit_photos_products_{selected_id}",
                        ),
                        "PACKAGING & CARTON DETAILS": st.file_uploader(
                            "Упаковка и маркировка", type=["jpg", "jpeg", "png", "webp"],
                            accept_multiple_files=True, key=f"edit_photos_packing_{selected_id}",
                        ),
                        "LABELS & BARCODES VERIFICATION": st.file_uploader(
                            "Стикеры и штрихкоды", type=["jpg", "jpeg", "png", "webp"],
                            accept_multiple_files=True, key=f"edit_photos_labels_{selected_id}",
                        ),
                        "DEFECT DETAILS & LOG": st.file_uploader(
                            "Дефекты", type=["jpg", "jpeg", "png", "webp"],
                            accept_multiple_files=True, key=f"edit_photos_defects_{selected_id}",
                        ),
                    }

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
                            "inspector_comments": edited_comments,
                            "comments": edited_comments,  # совместимость со старыми отчётами
                            "conclusion": edited_conclusion,
                        })

                        # Если добавлены фото, обновляем существующий Excel, не теряя исходные замеры и изображения.
                        if True:
                            selected_path_for_update = _storage_object_path(selected_report.get("excel_url"))
                            old_excel_bytes = download_excel_from_storage(selected_path_for_update) if selected_path_for_update else None
                            if not old_excel_bytes:
                                raise RuntimeError("Не удалось загрузить исходный Excel из Supabase; фотографии не сохранены.")
                            updated_for_excel = dict(updated_report_data)
                            updated_for_excel["inspection_date"] = normalized_date
                            updated_for_excel["inspector_id"] = st.session_state.auth_user.email
                            added_photos = []
                            for section_title, uploaded_files in edit_photo_uploads_by_section.items():
                                for item in (uploaded_files or []):
                                    added_photos.append((section_title, item.getvalue()))
                            refreshed_excel = append_photos_to_existing_excel(old_excel_bytes, added_photos, updated_for_excel)
                            upload_path = selected_path_for_update
                            upload_result = supabase.storage.from_(QC_REPORTS_BUCKET).upload(
                                upload_path,
                                refreshed_excel,
                                file_options={"content-type": EXCEL_MIME, "upsert": "true"},
                            )
                            if _storage_call_failed(upload_result):
                                raise RuntimeError("Supabase не подтвердил обновление Excel-файла.")
                            st.session_state[f"selected_excel_{selected_id}"] = refreshed_excel
                            st.session_state[f"selected_photos_{selected_id}"] = extract_images_from_excel(refreshed_excel)

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

            st.markdown("### Комментарии инспектора")
            comments = report_data.get("inspector_comments") or report_data.get("comments") or ""
            st.write(comments if comments else "—")
            st.markdown("### Заключение инспекции")
            conclusion = report_data.get("conclusion") or ""
            st.write(conclusion if conclusion else "—")

            # Загружаем Excel из Storage и показываем фотографии, встроенные в отчет.
            selected_path = _storage_object_path(selected_report.get("excel_url"))
            if selected_path:
                selected_cache_key = f"selected_excel_{selected_id}"
                selected_bytes = st.session_state.get(selected_cache_key)

                if selected_bytes is None:
                    with st.spinner("Загрузка фотографий отчета..."):
                        selected_bytes = download_excel_from_storage(selected_path)
                    if selected_bytes:
                        st.session_state[selected_cache_key] = selected_bytes

                if selected_bytes:
                    photos_cache_key = f"selected_photos_{selected_id}"
                    if photos_cache_key not in st.session_state:
                        st.session_state[photos_cache_key] = extract_images_from_excel(selected_bytes)

                    report_photos = st.session_state.get(photos_cache_key, [])
                    st.markdown("### Фотографии отчета")
                    if report_photos:
                        st.caption("Нажмите непосредственно на фотографию. В увеличенном просмотре доступны стрелки ← и →, а также клавиши клавиатуры.")
                        render_clickable_photo_gallery(report_photos)

                    else:
                        st.caption("В Excel этого отчета встроенные фотографии не найдены.")

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
                    st.warning("Не удалось загрузить Excel из облака, поэтому фотографии пока недоступны.")
            else:
                st.info("Для этого отчета Excel-файл в облачном хранилище не найден.")

        if not report_list and not selected_report:
            st.info("История отчетов пока пуста.")

    except Exception as e:
        st.warning("Не удалось загрузить архив. Проверьте подключение к Supabase и права доступа.")
        st.caption(str(e))


with tab_new:
    top_new_col, top_note_col = st.columns([1.5, 4])
    with top_new_col:
        st.button("＋ Создать новый отчёт", key="create_fresh_report", type="primary", use_container_width=True, on_click=_start_new_report)
    with top_note_col:
        st.caption("Новый отчёт создаёт чистый черновик. Уже сохранённые отчёты в архиве не изменяются.")
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
        report_no = col_b1.text_input("Report No.", key="draft_report_no")
        client_name = col_b2.text_input("Клиент / Заказчик")
        supplier_name = col_b3.text_input("Фабрика / Поставщик", value=extracted_data.get("supplier", ""))
        
        col_b4, col_b5, col_b6 = st.columns(3)
        order_no = col_b4.text_input("PO / Order No.", value=extracted_data.get("order_no", ""))
        item_name = col_b5.text_input("Наименование товара", value=extracted_data.get("item_name", ""))
        sku = col_b6.text_input("SKU / Артикул", value=extracted_data.get("sku", ""))
        inspection_date = st.date_input("Inspection Date / Дата инспекции", value=date.today(), key="draft_inspection_date")

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
                with st.form(key=f"sample_form_{s_num}_{selected_size}", enter_to_submit=False):
                    for pom in spec.get("poms", []):
                        p_code = pom.get("code", "")
                        p_desc = pom.get("description", "")
                        ideal_val = (pom.get("specs") or {}).get(selected_size)
                        t_min = pom.get("tol_minus")
                        t_plus = pom.get("tol_plus")

                        c1, c2, c3 = st.columns([1, 4, 3])
                        c1.write(f"**{p_code}**")
                        tol_txt = f"-{t_min if t_min is not None else '?'} / +{t_plus if t_plus is not None else '?'}"
                        c2.write(f"{p_desc} ({tol_txt})")
                        inputs[p_code] = {
                            "val": c3.text_input(
                                f"Замер {p_code}",
                                placeholder=str(ideal_val) if ideal_val is not None else "",
                                key=f"measurement_{s_num}_{selected_size}_{p_code}",
                                label_visibility="collapsed",
                            ),
                            "ideal": ideal_val,
                            "desc": p_desc,
                            "t_min": t_min,
                            "t_plus": t_plus
                        }
                    
                    if st.form_submit_button("✓ Проверить образец"):
                        sample_items = []
                        for code, item_meta in inputs.items():
                            raw_value = str(item_meta["val"] or "").strip()
                            reference_value = _to_number(item_meta["ideal"])
                            confirmed_reference = not raw_value and reference_value is not None
                            res = evaluate_measurement(
                                reference_value if confirmed_reference else raw_value,
                                item_meta["ideal"], item_meta["t_min"], item_meta["t_plus"]
                            )
                            sample_items.append({
                                "code": code,
                                "desc": item_meta["desc"],
                                "input_source": "confirmed_reference" if confirmed_reference else "manual",
                                "actual": res["actual"],
                                "diff": res["diff"],
                                "status": res["status"],
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
                        items = smp.get("items", [])
                        cnt = {k: sum(1 for it in items if it.get("status") == k) for k in MEASURE_STATUS_LABELS}
                        head = f"- Образец #{smp['sample_no']} ({smp['size']}): "
                        if items and cnt["PASS"] == len(items):
                            st.write(head + "✅ В норме")
                        else:
                            parts = [f"{MEASURE_STATUS_LABELS[k]}: {cnt[k]}" for k in ("FAIL", "NOT_MEASURED", "INVALID", "NO_TARGET") if cnt[k]]
                            st.write(head + "⚠️ " + "; ".join(parts))
                        bad_codes = [it.get("code") for it in items if it.get("status") == "INVALID"]
                        if bad_codes:
                            st.warning(
                                f"Образец #{smp['sample_no']}: не удалось распознать число в пунктах "
                                + ", ".join(str(c) for c in bad_codes)
                                + ". Вводите только число, например 72.5 или 72,5."
                            )

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
                preview_cols = st.columns(4)
                for idx, fl in enumerate(files):
                    with preview_cols[idx % 4]:
                        st.image(fl, use_container_width=True)
                    
                    photos_data[sec_key].append({
                        "bytes": fl.getvalue(),
                        "caption": f"{sec_title} — {idx + 1}",
                    })
            st.divider()

    # 4. Комментарии инспектора — кнопка обработки расположена ДО text_area,
    # чтобы не менять ключ session_state после создания виджета.
    with st.container(border=True):
        st.markdown("**4. Комментарии инспектора**")
        if st.button("✨ Отредактировать комментарии с ИИ", key="btn_ai_comments"):
            current_comments = st.session_state.get("comments_val", "")
            if current_comments.strip():
                with st.spinner("Редактирование текста с сохранением исходного языка..."):
                    language_guarded_text = (
                        "[INSTRUCTION: Correct and polish the text below. Preserve its original language exactly: "
                        "if the text is Russian, output Russian; if English, output English. Do not translate. "
                        "Return only the polished text.]\n\n" + current_comments
                    )
                    st.session_state.comments_val = ai_service.polish_qc_text(
                        language_guarded_text, field_type="comments"
                    )
                    st.rerun()
            else:
                st.info("Сначала введите комментарий.")
        raw_comments = st.text_area(
            "Комментарии инспектора",
            key="comments_val",
            placeholder="Рабочие заметки инспектора...",
            label_visibility="collapsed"
        )

    # 5. Заключение инспекции — тот же безопасный порядок.
    with st.container(border=True):
        st.markdown("**5. Заключение инспекции**")
        if st.button("✨ Отредактировать заключение с ИИ", key="btn_ai_conclusion"):
            current_conclusion = st.session_state.get("conclusion_val", "")
            if current_conclusion.strip():
                with st.spinner("Редактирование текста с сохранением исходного языка..."):
                    language_guarded_text = (
                        "[INSTRUCTION: Correct and polish the text below. Preserve its original language exactly: "
                        "if the text is Russian, output Russian; if English, output English. Do not translate. "
                        "Return only the polished text.]\n\n" + current_conclusion
                    )
                    st.session_state.conclusion_val = ai_service.polish_qc_text(
                        language_guarded_text, field_type="conclusion"
                    )
                    st.rerun()
            else:
                st.info("Сначала введите заключение.")
        final_comments = st.text_area(
            "Заключение инспекции",
            key="conclusion_val",
            placeholder="Итоговое заключение по результатам проверки...",
            label_visibility="collapsed"
        )

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
            "inspection_date": inspection_date.isoformat(),
            "result_status": result_status,
            "inspector_comments": raw_comments,
            "comments": raw_comments,  # backward compatibility with older report records
            "conclusion": final_comments
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
                        "inspection_date": inspection_date.isoformat(),
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

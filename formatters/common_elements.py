# common_elements.py
import time
import traceback
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from utils import set_paragraph_format, set_run_format, add_run_with_format, add_paragraph_with_text
from config import (
    FONT_NAME, FONT_SIZE_HEADER_13, FONT_SIZE_MEDIUM_14, FONT_SIZE_MEDIUM_13,
    FONT_SIZE_SIGN_AUTH_14, FONT_SIZE_SIGN_NAME_14,
    FONT_SIZE_RECIPIENT_LABEL_12, FONT_SIZE_RECIPIENT_LIST_11,
    FONT_SIZE_OTHER_11
)

# --- Các hàm add_... với try-except chi tiết hơn ---

def add_quoc_hieu_tieu_ngu(table_cell):
    print("  common: Adding Quoc Hieu Tieu Ngu...")
    try:
        paragraph_qh = table_cell.add_paragraph()
        set_paragraph_format(paragraph_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(paragraph_qh, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", size=FONT_SIZE_HEADER_13, bold=True)

        paragraph_tn = table_cell.add_paragraph()
        set_paragraph_format(paragraph_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(paragraph_tn, "Độc lập - Tự do - Hạnh phúc", size=FONT_SIZE_MEDIUM_14, bold=True)

        paragraph_line = table_cell.add_paragraph()
        set_paragraph_format(paragraph_line, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(paragraph_line, "_______________", size=FONT_SIZE_MEDIUM_14, bold=True)
        print("  common: Added Quoc Hieu Tieu Ngu OK.")
    except Exception as e:
        print(f"  ERROR in add_quoc_hieu_tieu_ngu: {e}")
        print(traceback.format_exc())
        try: table_cell.add_paragraph(f"[Lỗi QH-TN: {e}]")
        except: pass # Bỏ qua nếu không thêm được vào cell

def add_ten_co_quan_ban_hanh(table_cell, ten_co_quan_chu_quan, ten_co_quan_ban_hanh):
    print("  common: Adding Ten Co Quan...")
    try:
        if ten_co_quan_chu_quan:
            paragraph_cqcq = table_cell.add_paragraph()
            set_paragraph_format(paragraph_cqcq, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
            add_run_with_format(paragraph_cqcq, ten_co_quan_chu_quan.upper(), size=FONT_SIZE_HEADER_13, bold=False)

        paragraph_cqbh = table_cell.add_paragraph()
        set_paragraph_format(paragraph_cqbh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(paragraph_cqbh, ten_co_quan_ban_hanh.upper(), size=FONT_SIZE_HEADER_13, bold=True)

        paragraph_line = table_cell.add_paragraph()
        set_paragraph_format(paragraph_line, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(paragraph_line, "________", size=FONT_SIZE_HEADER_13, bold=True)
        print("  common: Added Ten Co Quan OK.")
    except Exception as e:
        print(f"  ERROR in add_ten_co_quan_ban_hanh: {e}")
        print(traceback.format_exc())
        try: table_cell.add_paragraph(f"[Lỗi Ten CQ: {e}]")
        except: pass

def add_so_ky_hieu(table_cell, so_van_ban, ky_hieu_van_ban):
    print("  common: Adding So Ky Hieu...")
    try:
        paragraph_skh = table_cell.add_paragraph()
        set_paragraph_format(paragraph_skh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(paragraph_skh, f"Số: {so_van_ban}/", size=FONT_SIZE_MEDIUM_13, bold=False)
        add_run_with_format(paragraph_skh, ky_hieu_van_ban.upper(), size=FONT_SIZE_MEDIUM_13, bold=False)
        print("  common: Added So Ky Hieu OK.")
    except Exception as e:
        print(f"  ERROR in add_so_ky_hieu: {e}")
        print(traceback.format_exc())
        try: table_cell.add_paragraph(f"[Lỗi SKH: {e}]")
        except: pass

def add_dia_danh_thoi_gian(table_cell, dia_danh, ngay, thang, nam):
    print("  common: Adding Dia Danh Thoi Gian...")
    try:
        paragraph_ddtg = table_cell.add_paragraph()
        set_paragraph_format(paragraph_ddtg, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        thoi_gian_str = f"ngày {ngay:02d} tháng {thang:02d} năm {nam}"
        add_run_with_format(paragraph_ddtg, f"{dia_danh}, {thoi_gian_str}", size=FONT_SIZE_MEDIUM_13, italic=True)
        print("  common: Added Dia Danh Thoi Gian OK.")
    except Exception as e:
        print(f"  ERROR in add_dia_danh_thoi_gian: {e}")
        print(traceback.format_exc())
        try: table_cell.add_paragraph(f"[Lỗi DDTG: {e}]")
        except: pass

def add_signature_block(
    document, authority_signer=None, signer_title="", signer_name="",
    signer_note=None, alignment=WD_ALIGN_PARAGRAPH.RIGHT, use_table=True
):
    print("  common: Adding Signature Block...")
    try:
        if use_table:
            # ... (logic tạo bảng và cell_sig như cũ) ...
            sig_table = document.add_table(rows=1, cols=2)
            sig_table.autofit = False
            sig_table.allow_autofit = False
            sig_table.columns[0].width = Inches(3.0)
            sig_table.columns[1].width = Inches(3.3)
            cell_sig = sig_table.cell(0, 1)
            cell_sig._element.clear_content()
            paragraph_container = cell_sig
            sig_align = WD_ALIGN_PARAGRAPH.CENTER
        else:
            paragraph_container = document
            sig_align = alignment

        if authority_signer:
            paragraph_auth = paragraph_container.add_paragraph()
            set_paragraph_format(paragraph_auth, alignment=sig_align, space_after=Pt(0))
            add_run_with_format(paragraph_auth, authority_signer.upper(), size=FONT_SIZE_SIGN_AUTH_14, bold=True)

        paragraph_title = paragraph_container.add_paragraph()
        set_paragraph_format(paragraph_title, alignment=sig_align, space_after=Pt(0))
        add_run_with_format(paragraph_title, signer_title.upper(), size=FONT_SIZE_SIGN_AUTH_14, bold=True)

        if signer_note:
            paragraph_note = paragraph_container.add_paragraph()
            set_paragraph_format(paragraph_note, alignment=sig_align, space_after=Pt(0))
            add_run_with_format(paragraph_note, signer_note, size=FONT_SIZE_OTHER_11, italic=True)

        paragraph_space = paragraph_container.add_paragraph("\n\n\n\n")
        set_paragraph_format(paragraph_space, alignment=sig_align, space_after=Pt(0))

        paragraph_name = paragraph_container.add_paragraph()
        set_paragraph_format(paragraph_name, alignment=sig_align, space_after=Pt(0))
        add_run_with_format(paragraph_name, signer_name, size=FONT_SIZE_SIGN_NAME_14, bold=True)
        print("  common: Added Signature Block OK.")
    except Exception as e:
        print(f"  ERROR in add_signature_block: {e}")
        print(traceback.format_exc())
        try:
            # Thêm lỗi vào doc để dễ debug
            container = cell_sig if use_table else document
            container.add_paragraph(f"[Lỗi Chữ ký: {e}]")
        except: pass


def add_recipient_list(document, recipients):
    print("  common: Adding Recipient List...")
    try:
        paragraph_label = document.add_paragraph()
        set_paragraph_format(paragraph_label, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(6), space_after=Pt(0))
        add_run_with_format(paragraph_label, "Nơi nhận:", size=FONT_SIZE_RECIPIENT_LABEL_12, bold=True, italic=True)

        if not recipients:
            recipients = ["- Như Điều ...;", "- Lưu: VT, ...;"]

        for recipient in recipients:
            paragraph_rec = document.add_paragraph()
            set_paragraph_format(
                paragraph_rec, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                left_indent=Cm(0.7), first_line_indent=Cm(-0.7),
                space_before=Pt(0), space_after=Pt(0), line_spacing=1.0
            )
            add_run_with_format(paragraph_rec, recipient, size=FONT_SIZE_RECIPIENT_LIST_11, bold=False)
        print("  common: Added Recipient List OK.")
    except Exception as e:
        print(f"  ERROR in add_recipient_list: {e}")
        print(traceback.format_exc())
        try: document.add_paragraph(f"[Lỗi Nơi nhận: {e}]")
        except: pass
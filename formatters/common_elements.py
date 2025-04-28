# common_elements.py
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from utils import set_paragraph_format, set_run_format, add_run_with_format, add_paragraph_with_text
# Import các hằng số đã chuẩn hóa từ config
from config import (
    FONT_NAME,
    FONT_SIZE_HEADER_13, FONT_SIZE_MEDIUM_14, FONT_SIZE_MEDIUM_13,
    FONT_SIZE_SIGN_AUTH_14, FONT_SIZE_SIGN_NAME_14,
    FONT_SIZE_RECIPIENT_LABEL_12, FONT_SIZE_RECIPIENT_LIST_11,
    FONT_SIZE_OTHER_11
)

def add_quoc_hieu_tieu_ngu(table_cell):
    """Thêm Quốc hiệu và Tiêu ngữ vào ô bảng (Ô số 1)."""
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
    except Exception as e:
        print(f"Error adding Quoc Hieu Tieu Ngu: {e}")
        table_cell.add_paragraph(f"[Error QH-TN: {e}]") # Ghi lỗi vào cell

def add_ten_co_quan_ban_hanh(table_cell, ten_co_quan_chu_quan, ten_co_quan_ban_hanh):
    """Thêm Tên cơ quan chủ quản (nếu có) và Tên cơ quan ban hành (Ô số 2)."""
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
    except Exception as e:
        print(f"Error adding Ten Co Quan: {e}")
        table_cell.add_paragraph(f"[Error Ten CQ: {e}]")

def add_so_ky_hieu(table_cell, so_van_ban, ky_hieu_van_ban):
    """Thêm Số, ký hiệu văn bản vào ô bảng (Ô số 3)."""
    try:
        paragraph_skh = table_cell.add_paragraph()
        set_paragraph_format(paragraph_skh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(paragraph_skh, f"Số: {so_van_ban}/", size=FONT_SIZE_MEDIUM_13, bold=False)
        add_run_with_format(paragraph_skh, ky_hieu_van_ban.upper(), size=FONT_SIZE_MEDIUM_13, bold=False)
    except Exception as e:
        print(f"Error adding So Ky Hieu: {e}")
        table_cell.add_paragraph(f"[Error SKH: {e}]")

def add_dia_danh_thoi_gian(table_cell, dia_danh, ngay, thang, nam):
    """Thêm Địa danh, thời gian ban hành vào ô bảng (Ô số 4)."""
    try:
        paragraph_ddtg = table_cell.add_paragraph()
        # Căn giữa trong ô phải của bảng header
        set_paragraph_format(paragraph_ddtg, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        thoi_gian_str = f"ngày {ngay:02d} tháng {thang:02d} năm {nam}"
        add_run_with_format(paragraph_ddtg, f"{dia_danh}, {thoi_gian_str}", size=FONT_SIZE_MEDIUM_13, italic=True)
    except Exception as e:
        print(f"Error adding Dia Danh Thoi Gian: {e}")
        table_cell.add_paragraph(f"[Error DDTG: {e}]")


def add_signature_block(
    document,
    authority_signer=None,
    signer_title="",
    signer_name="",
    signer_note=None,
    alignment=WD_ALIGN_PARAGRAPH.RIGHT, # Mặc định căn phải cho HC
    use_table=True
):
    """Thêm khối chữ ký (Ô số 7a, 7b, 7c)."""
    try:
        if use_table:
            sig_table = document.add_table(rows=1, cols=2)
            sig_table.autofit = False
            sig_table.allow_autofit = False
            sig_table.columns[0].width = Inches(3.0)
            sig_table.columns[1].width = Inches(3.3)
            cell_sig = sig_table.cell(0, 1)
            cell_sig._element.clear_content()
            paragraph_container = cell_sig
            sig_align = WD_ALIGN_PARAGRAPH.CENTER # Căn giữa trong ô phải
        else:
            paragraph_container = document
            sig_align = alignment # Dùng alignment truyền vào nếu ko dùng table

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
    except Exception as e:
        print(f"Error adding Signature Block: {e}")
        # Thêm lỗi vào doc để dễ debug nếu dùng table
        if use_table:
             cell_sig.add_paragraph(f"[Error Signature: {e}]")
        else:
             document.add_paragraph(f"[Error Signature: {e}]")


def add_recipient_list(document, recipients):
    """Thêm danh sách nơi nhận (Ô số 9b)."""
    try:
        paragraph_label = document.add_paragraph()
        set_paragraph_format(paragraph_label, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(6), space_after=Pt(0)) # Giảm space_before
        add_run_with_format(paragraph_label, "Nơi nhận:", size=FONT_SIZE_RECIPIENT_LABEL_12, bold=True, italic=True)

        if not recipients:
            recipients = ["- Như Điều ...;", "- Lưu: VT, ...;"]

        for recipient in recipients:
            paragraph_rec = document.add_paragraph()
            set_paragraph_format(
                paragraph_rec,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                left_indent=Cm(0.7),
                first_line_indent=Cm(-0.7),
                space_before=Pt(0),
                space_after=Pt(0),
                line_spacing=1.0
            )
            add_run_with_format(paragraph_rec, recipient, size=FONT_SIZE_RECIPIENT_LIST_11, bold=False)
    except Exception as e:
        print(f"Error adding Recipient List: {e}")
        document.add_paragraph(f"[Error Noi Nhan: {e}]")
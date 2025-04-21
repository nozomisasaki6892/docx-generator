# formatters/quy_dinh.py
import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Quy định...")
    title = data.get("title", "Quy định về ABC")
    body = data.get("body", "Nội dung quy định...")
    # Quy định thường ban hành kèm Quyết định, nên header có thể của CQ ban hành QĐ
    issuing_authority = data.get("issuing_org", "CƠ QUAN BAN HÀNH").upper()

    # Giả sử Quy định là văn bản độc lập hoặc phần chính
    # Header có thể không cần nếu là Phụ lục kèm Quyết định
    # format_basic_header(document, data, "QuyDinh") # Tạm bỏ qua header

    # Tên loại
    p_tenloai = document.add_paragraph("QUY ĐỊNH")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "QUY ĐỊNH", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # Tiêu đề của Quy định
    quy_dinh_title = title.replace("Quy định", "").strip()
    p_title = document.add_paragraph(quy_dinh_title)
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_title, quy_dinh_title, size=Pt(14), bold=True)

    # Dòng "Ban hành kèm theo Quyết định số..." (Nếu có)
    attached_decision = data.get("attached_decision", None)
    if attached_decision:
        p_attach = document.add_paragraph(f"(Ban hành kèm theo {attached_decision})")
        set_paragraph_format(p_attach, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
        add_run_with_format(p_attach, f"(Ban hành kèm theo {attached_decision})", size=FONT_SIZE_DEFAULT, italic=True)

    # Nội dung (Thường có Chương, Mục, Điều, Khoản, Điểm)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if not stripped_line: continue

        p = document.add_paragraph()
        is_chuong = stripped_line.upper().startswith("CHƯƠNG")
        is_muc = re.match(r'^(MỤC\s+\d+)\.?\s+', stripped_line.upper())
        is_dieu = stripped_line.upper().startswith("ĐIỀU")
        is_khoan = re.match(r'^\d+\.\s+', stripped_line)
        is_diem = re.match(r'^[a-z]\)\s+', stripped_line)

        align = WD_ALIGN_PARAGRAPH.JUSTIFY
        left_indent = Cm(0)
        first_indent = FIRST_LINE_INDENT
        is_bold = False
        size = FONT_SIZE_DEFAULT

        if is_chuong:
            align = WD_ALIGN_PARAGRAPH.CENTER
            first_indent = Cm(0)
            is_bold = True
            size = FONT_SIZE_DEFAULT # Hoặc 14
            # Thêm space before/after
        elif is_muc:
             align = WD_ALIGN_PARAGRAPH.CENTER
             first_indent = Cm(0)
             is_bold = True
             size = FONT_SIZE_DEFAULT
        elif is_dieu:
             align = WD_ALIGN_PARAGRAPH.LEFT # Điều căn trái
             first_indent = Cm(0)
             is_bold = True
        elif is_khoan:
             left_indent = Cm(0.5)
             first_indent = Cm(0)
        elif is_diem:
             left_indent = Cm(1.0)
             first_indent = Cm(0)

        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_after=Pt(6))
        add_run_with_format(p, stripped_line, size=size, bold=is_bold)

    # Quy định thường không có chữ ký, nơi nhận riêng nếu ban hành kèm QĐ
    # Nếu là quy định độc lập thì cần thêm
    # format_signature_block(document, data)
    # format_recipient_list(document, data)

    print("Định dạng Quy định hoàn tất.")
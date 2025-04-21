# formatters/huong_dan_hs.py
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
    print("Bắt đầu định dạng Hướng dẫn hồ sơ...")
    title = data.get("title", "Hướng dẫn hồ sơ đăng ký dự thi / nhập học")
    body = data.get("body", "Nội dung hướng dẫn...")
    issuing_org = data.get("issuing_org", "TÊN TRƯỜNG/ĐƠN VỊ").upper()

    # 1. Header của trường
    data['issuing_org'] = issuing_org
    format_basic_header(document, data, "HuongDanHS")

    # 2. Tên loại
    p_tenloai = document.add_paragraph("HƯỚNG DẪN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "HƯỚNG DẪN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Tiêu đề hướng dẫn
    hd_title = title.replace("Hướng dẫn", "").strip()
    p_title = document.add_paragraph(hd_title)
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, hd_title, size=Pt(14), bold=True)

    # 4. Nội dung (Các bước I, 1, a, -)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            # Copy logic nhận diện mục từ huong_dan.py hoặc ke_hoach.py
            left_indent_val = Cm(0)
            first_indent_val = FIRST_LINE_INDENT
            is_bold_run = False
            align = WD_ALIGN_PARAGRAPH.JUSTIFY

            is_roman = re.match(r'^[IVXLCDM]+\.\s+', stripped_line)
            is_arabic = re.match(r'^\d+\.\s+', stripped_line)
            is_alpha = re.match(r'^[a-z]\)\s+', stripped_line)
            is_dash = stripped_line.startswith('-')

            if is_roman:
                is_bold_run = True
                align = WD_ALIGN_PARAGRAPH.LEFT
                first_indent_val = Cm(0)
                set_paragraph_format(p, alignment=align, left_indent=Cm(0), first_line_indent=Cm(0), space_before=Pt(12), space_after=Pt(6))
            elif is_arabic:
                left_indent_val = Cm(0.5)
                first_indent_val = Cm(0)
                is_bold_run = True # Mục 1, 2 đậm
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent_val, first_line_indent=Cm(0), space_after=Pt(6))
            elif is_alpha:
                left_indent_val = Cm(1.0)
                first_indent_val = Cm(0)
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent_val, first_line_indent=Cm(0), space_after=Pt(6))
            elif is_dash:
                left_indent_val = Cm(1.5)
                first_indent_val = Cm(0)
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent_val, first_line_indent=Cm(0), space_after=Pt(6))
            else:
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0), first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))

            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_bold_run)


    # 5. Chữ ký (Người lập hướng dẫn hoặc lãnh đạo đơn vị)
    if not data.get('signer_title'): data['signer_title'] = "TRƯỞNG PHÒNG ĐÀO TẠO" # Ví dụ
    format_signature_block(document, data)

    # 6. Nơi nhận (nếu cần)
    # format_recipient_list(document, data)

    print("Định dạng Hướng dẫn hồ sơ hoàn tất.")
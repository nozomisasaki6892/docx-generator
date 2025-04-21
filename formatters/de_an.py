# formatters/de_an.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Đề án...")
    title = data.get("title", "Đề án ABC")
    body = data.get("body", "Nội dung đề án...")

    # Đề án thường trình cấp trên nên có Header chuẩn
    format_basic_header(document, data, "DeAn")

    p_tenloai = document.add_paragraph("ĐỀ ÁN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "ĐỀ ÁN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # Tiêu đề của Đề án
    de_an_title = title.replace("Đề án", "").strip()
    p_title = document.add_paragraph(de_an_title)
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, de_an_title, size=Pt(14), bold=True)

    # Nội dung (Cấu trúc I, 1, a, - hoặc tương tự Kế hoạch, Chương trình)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            left_indent_val = Cm(0)
            first_indent_val = FIRST_LINE_INDENT
            is_bold_run = False
            align = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Nhận diện các mục tiêu đề chính của Đề án
            is_main_heading = any(stripped_line.upper().startswith(h) for h in ["I.", "II.", "III.", "IV.", "V.", "PHẦN", "MỤC LỤC", "SỰ CẦN THIẾT", "MỤC TIÊU", "PHẠM VI", "NỘI DUNG", "GIẢI PHÁP", "KINH PHÍ", "TỔ CHỨC THỰC HIỆN", "HIỆU QUẢ"])

            is_roman = re.match(r'^[IVXLCDM]+\.\s+', stripped_line)
            is_arabic = re.match(r'^\d+\.\s+', stripped_line)
            is_alpha = re.match(r'^[a-z]\)\s+', stripped_line)
            is_dash = stripped_line.startswith('-')

            if is_main_heading or is_roman:
                is_bold_run = True
                align = WD_ALIGN_PARAGRAPH.LEFT # Các mục lớn thường căn trái
                first_indent_val = Cm(0)
                # Có thể tăng khoảng cách trước các mục lớn
                set_paragraph_format(p, alignment=align, left_indent=Cm(0), first_line_indent=Cm(0), space_before=Pt(12), space_after=Pt(6))
            elif is_arabic:
                 left_indent_val = Cm(0.5)
                 first_indent_val = Cm(0)
                 is_bold_run = True # Mục cấp 2 cũng có thể đậm
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


    # Chữ ký (Thường là người đứng đầu cơ quan chủ trì Đề án)
    format_signature_block(document, data)

    # Nơi nhận (Thường gửi cấp trên phê duyệt và các đơn vị phối hợp)
    format_recipient_list(document, data)

    print("Định dạng Đề án hoàn tất.")
# formatters/du_an.py
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
    print("Bắt đầu định dạng Dự án...")
    title = data.get("title", "Dự án đầu tư xây dựng ABC")
    body = data.get("body", "Nội dung dự án...")

    # Dự án thường trình duyệt nên có Header
    format_basic_header(document, data, "DuAn")

    p_tenloai = document.add_paragraph("DỰ ÁN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "DỰ ÁN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # Tên của Dự án
    da_title = title.replace("Dự án", "").strip()
    p_title = document.add_paragraph(da_title)
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, da_title, size=Pt(14), bold=True)

    # Nội dung cấu trúc tương tự Đề án, thường rất chi tiết
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            left_indent_val = Cm(0)
            first_indent_val = FIRST_LINE_INDENT
            is_bold_run = False
            align = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Nhận diện các mục tiêu đề chính của Dự án
            is_main_heading = any(stripped_line.upper().startswith(h) for h in ["I.", "II.", "III.", "IV.", "V.", "PHẦN", "MỤC LỤC", "SỰ CẦN THIẾT", "MỤC TIÊU", "ĐỊA ĐIỂM", "QUY MÔ", "GIẢI PHÁP THIẾT KẾ", "TIẾN ĐỘ", "KINH PHÍ", "TỔNG MỨC ĐẦU TƯ", "NGUỒN VỐN", "HIỆU QUẢ", "TỔ CHỨC QUẢN LÝ"])

            is_roman = re.match(r'^[IVXLCDM]+\.\s+', stripped_line)
            is_arabic = re.match(r'^\d+\.\s+', stripped_line)
            is_alpha = re.match(r'^[a-z]\)\s+', stripped_line)
            is_dash = stripped_line.startswith('-')

            if is_main_heading or is_roman:
                is_bold_run = True
                align = WD_ALIGN_PARAGRAPH.LEFT
                first_indent_val = Cm(0)
                set_paragraph_format(p, alignment=align, left_indent=Cm(0), first_line_indent=Cm(0), space_before=Pt(12), space_after=Pt(6))
            elif is_arabic:
                left_indent_val = Cm(0.5)
                first_indent_val = Cm(0)
                is_bold_run = True
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

    # Chữ ký của chủ đầu tư hoặc cơ quan lập dự án
    format_signature_block(document, data)
    # Nơi nhận là cấp phê duyệt, các sở ban ngành liên quan
    format_recipient_list(document, data)

    print("Định dạng Dự án hoàn tất.")
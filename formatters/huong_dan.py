# formatters/huong_dan.py
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
    print("Bắt đầu định dạng Hướng dẫn...")
    title = data.get("title", "Hướng dẫn thực hiện ABC")
    body = data.get("body", "Nội dung hướng dẫn...")

    format_basic_header(document, data, "HuongDan")

    p_tenloai = document.add_paragraph("HƯỚNG DẪN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "HƯỚNG DẪN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    trich_yeu_text = title.replace("Hướng dẫn", "").strip()
    p_trichyeu = document.add_paragraph(trich_yeu_text) # Hướng dẫn thường ko có "V/v"
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_trichyeu, trich_yeu_text, size=Pt(14), bold=True)

    # Nội dung thường có các phần I, II, 1, 2, a, b
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            left_indent_val = Cm(0)
            first_indent_val = FIRST_LINE_INDENT
            is_bold_run = False
            align = WD_ALIGN_PARAGRAPH.JUSTIFY

            is_roman = re.match(r'^[IVXLCDM]+\.\s+', stripped_line) # I. II. ...
            is_arabic = re.match(r'^\d+\.\s+', stripped_line)      # 1. 2. ...
            is_alpha = re.match(r'^[a-z]\)\s+', stripped_line)     # a) b) ...
            is_dash = stripped_line.startswith('-')              # - ...

            if is_roman:
                is_bold_run = True
                align = WD_ALIGN_PARAGRAPH.LEFT
                first_indent_val = Cm(0)
            elif is_arabic:
                 left_indent_val = Cm(0.5)
                 first_indent_val = Cm(0)
                 # Có thể in đậm mục 1, 2 nếu muốn
            elif is_alpha:
                 left_indent_val = Cm(1.0)
                 first_indent_val = Cm(0)
            elif is_dash:
                 left_indent_val = Cm(1.5)
                 first_indent_val = Cm(0)

            set_paragraph_format(p, alignment=align, left_indent=left_indent_val, first_line_indent=first_indent_val, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_bold_run)

    format_signature_block(document, data)
    format_recipient_list(document, data)

    print("Định dạng Hướng dẫn hoàn tất.")
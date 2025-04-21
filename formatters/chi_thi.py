# formatters/chi_thi.py
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
    print("Bắt đầu định dạng Chỉ thị...")
    title = data.get("title", "Chỉ thị về việc ABC")
    body = data.get("body", "Nội dung chỉ thị...")

    format_basic_header(document, data, "ChiThi")

    p_tenloai = document.add_paragraph("CHỈ THỊ")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "CHỈ THỊ", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    trich_yeu_text = title.replace("Chỉ thị", "").strip()
    # Chỉ thị thường không có "Về việc" ở trích yếu
    p_trichyeu = document.add_paragraph(trich_yeu_text)
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_trichyeu, trich_yeu_text, size=Pt(14), bold=True)

    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            is_numbered_item = re.match(r'^\d+\.\s+', stripped_line)

            if is_numbered_item:
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0.5), first_line_indent=Cm(0), space_after=Pt(6))
                match = re.match(r'^(\d+\.)(\s+.*)', stripped_line)
                if match:
                    add_run_with_format(p, match.group(1), size=FONT_SIZE_DEFAULT, bold=True)
                    add_run_with_format(p, match.group(2), size=FONT_SIZE_DEFAULT)
                else:
                    add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=True)
            else:
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0), first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))
                add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    format_signature_block(document, data)
    format_recipient_list(document, data)

    print("Định dạng Chỉ thị hoàn tất.")
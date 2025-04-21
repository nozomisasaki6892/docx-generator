# formatters/thong_bao.py
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
    print("Bắt đầu định dạng Thông báo...")
    title = data.get("title", "Thông báo về việc ABC")
    body = data.get("body", "Nội dung thông báo...")

    format_basic_header(document, data, "ThongBao")

    p_tenloai = document.add_paragraph("THÔNG BÁO")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "THÔNG BÁO", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    trich_yeu_text = title.replace("Thông báo", "").strip()
    p_trichyeu = document.add_paragraph(f"Về việc {trich_yeu_text}")
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_trichyeu, f"Về việc {trich_yeu_text}", size=Pt(14), bold=True)

    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            is_kinh_gui = stripped_line.upper().startswith("KÍNH GỬI")
            align = WD_ALIGN_PARAGRAPH.JUSTIFY
            indent = FIRST_LINE_INDENT if is_kinh_gui else Cm(0)
            first_indent_val = Cm(0) if is_kinh_gui else FIRST_LINE_INDENT

            set_paragraph_format(p, alignment=align, left_indent=indent if is_kinh_gui else Cm(0), first_line_indent=first_indent_val, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_kinh_gui)

    format_signature_block(document, data)
    format_recipient_list(document, data)

    print("Định dạng Thông báo hoàn tất.")
# formatters/cong_van.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list

from config import FONT_SIZE_DEFAULT, FONT_SIZE_SMALL, FIRST_LINE_INDENT, FONT_SIZE_VV

def format(document, data):
    print("Đang định dạng Công văn...")
    title = data.get("title", "Về việc ABC")
    body = data.get("body", "Nội dung công văn...")
    recipients_to = data.get("recipients_to", "Kính gửi: [Tên đơn vị/cá nhân]")

    format_basic_header(document, data, "CongVan")

    trich_yeu_text = title
    if title.lower().startswith("về việc"):
        trich_yeu_text = title[len("về việc"):].strip()
    elif title.lower().startswith("v/v"):
         trich_yeu_text = title[len("v/v"):].strip()

    p_vv = document.add_paragraph()
    set_paragraph_format(p_vv, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(6))
    add_run_with_format(p_vv, f"V/v: {trich_yeu_text}", size=FONT_SIZE_VV)

    p_kg = document.add_paragraph()
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=Cm(0), first_line_indent=FIRST_LINE_INDENT, space_before=Pt(6), space_after=Pt(6)) # Kính gửi thẳng lề nội dung
    add_run_with_format(p_kg, recipients_to, size=FONT_SIZE_DEFAULT, bold=True)

    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    format_signature_block(document, data)
    format_recipient_list(document, data)

    print("Định dạng Công văn hoàn tất.")
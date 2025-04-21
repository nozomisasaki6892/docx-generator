# formatters/thong_cao.py
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
    print("Bắt đầu định dạng Thông cáo...")
    title = data.get("title", "Thông cáo báo chí về sự kiện ABC")
    body = data.get("body", "Nội dung thông cáo...")

    # Thông cáo thường do cơ quan, tổ chức phát hành nên có header
    format_basic_header(document, data, "ThongCao")

    p_tenloai = document.add_paragraph("THÔNG CÁO")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "THÔNG CÁO", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # Tiêu đề/Trích yếu của Thông cáo
    tc_title = title.replace("Thông cáo", "").strip()
    p_title = document.add_paragraph(tc_title)
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, tc_title, size=Pt(14), bold=True)

    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    # Thông cáo thường do người đứng đầu ký
    format_signature_block(document, data)

    # Nơi nhận có thể là các cơ quan báo chí, đơn vị liên quan
    format_recipient_list(document, data)

    print("Định dạng Thông cáo hoàn tất.")
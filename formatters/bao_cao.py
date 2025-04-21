# formatters/bao_cao.py
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
    print("Bắt đầu định dạng Báo cáo...")
    title = data.get("title", "Báo cáo về việc ABC")
    body = data.get("body", "Nội dung báo cáo...")
    recipients_to = data.get("recipients_to", None) # "Kính gửi:..." là tùy chọn
    issuing_org = data.get("issuing_org", "TÊN ĐƠN VỊ BÁO CÁO").upper()

    # 1. Header
    data['issuing_org'] = issuing_org
    format_basic_header(document, data, "BaoCao") # Header căn trái CQBH

    # 2. Tên loại BÁO CÁO
    p_tenloai = document.add_paragraph("BÁO CÁO")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "BÁO CÁO", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Trích yếu
    bc_title = title.replace("Báo cáo", "").strip()
    p_title = document.add_paragraph(f"Về việc {bc_title}")
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, f"Về việc {bc_title}", size=Pt(14), bold=True)

    # 4. Kính gửi (Nếu có)
    if recipients_to:
        p_kg = document.add_paragraph(recipients_to)
        set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
        add_run_with_format(p_kg, recipients_to, size=FONT_SIZE_DEFAULT, bold=True)

    # 5. Nội dung Báo cáo
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            # Nhận diện mục La Mã, Ả Rập...
            is_roman = re.match(r'^[IVXLCDM]+\.\s+', stripped_line)
            is_arabic = re.match(r'^\d+\.\s+', stripped_line)
            is_alpha = re.match(r'^[a-z]\)\s+', stripped_line)
            is_dash = stripped_line.startswith('-')
            is_main_heading = is_roman or is_arabic

            align = WD_ALIGN_PARAGRAPH.JUSTIFY
            left_indent = Cm(0)
            first_indent = FIRST_LINE_INDENT
            is_bold = False

            if is_roman:
                align = WD_ALIGN_PARAGRAPH.LEFT
                first_indent = Cm(0)
                is_bold = True
                set_paragraph_format(p, alignment=align, left_indent=Cm(0), first_line_indent=Cm(0), space_before=Pt(12), space_after=Pt(6))
            elif is_arabic:
                left_indent = Cm(0.5)
                first_indent = Cm(0)
                is_bold = True
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent, first_line_indent=first_indent, space_after=Pt(6))
            elif is_alpha:
                left_indent = Cm(1.0)
                first_indent = Cm(0)
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent, first_line_indent=first_indent, space_after=Pt(6))
            elif is_dash:
                left_indent = Cm(1.5)
                first_indent = Cm(0)
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent, first_line_indent=first_indent, space_after=Pt(6))
            else:
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0), first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))

            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_bold)

    # 6. Chữ ký
    format_signature_block(document, data)

    # 7. Nơi nhận (Tùy chọn)
    if data.get('recipients'):
         format_recipient_list(document, data)

    print("Định dạng Báo cáo hoàn tất.")
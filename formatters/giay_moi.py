# formatters/giay_moi.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_VV

def format(document, data):
    print("Bắt đầu định dạng Giấy mời...")
    title = data.get("title", "Giấy mời họp/tham dự sự kiện ABC")
    body = data.get("body", "Nội dung giấy mời...")
    recipients_to = data.get("recipients_to", "Kính gửi: Ông/Bà [Tên người được mời]")
    event_subject = data.get("event_subject", title.replace("Giấy mời", "").strip()) # Chủ đề sự kiện
    issuing_org = data.get("issuing_org", "TÊN ĐƠN VỊ MỜI").upper()

    # 1. Header
    data['issuing_org'] = issuing_org
    format_basic_header(document, data, "GiayMoi")

    # 2. V/v (Nếu có, căn trái dưới số KH)
    p_vv = document.add_paragraph()
    set_paragraph_format(p_vv, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(6))
    add_run_with_format(p_vv, f"V/v: {event_subject}", size=FONT_SIZE_VV) # Size 12

    # 3. Tên loại GIẤY MỜI
    p_tenloai = document.add_paragraph("GIẤY MỜI")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_tenloai, "GIẤY MỜI", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Kính gửi
    p_kg = document.add_paragraph(recipients_to)
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_kg, recipients_to, size=FONT_SIZE_DEFAULT, bold=True)

    # 5. Nội dung mời
    p_intro = document.add_paragraph()
    set_paragraph_format(p_intro, first_line_indent=FIRST_LINE_INDENT)
    add_run_with_format(p_intro, f"{issuing_org} trân trọng kính mời Ông/Bà tới dự:") # Hoặc Đến tham dự...

    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            # Nhận diện các mục thông tin (Thời gian, Địa điểm, Nội dung...)
            is_info_heading = any(stripped_line.startswith(h) for h in ["Thời gian:", "Địa điểm:", "Nội dung:", "Thành phần:", "Chủ trì:", "Chương trình:"])
            left_indent = Cm(0.5) if is_info_heading else Cm(0)
            first_indent = Cm(0) if is_info_heading else FIRST_LINE_INDENT

            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_after=Pt(6))
            # Có thể tách và in đậm phần label (Thời gian:, Địa điểm:)
            match = re.match(r'^([\w\s]+:)(.*)', stripped_line)
            if match and is_info_heading:
                add_run_with_format(p, match.group(1), bold=True)
                add_run_with_format(p, match.group(2))
            else:
                add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    # 6. Lời kết
    p_closing = document.add_paragraph("Rất mong Ông/Bà sắp xếp thời gian tham dự.")
    set_paragraph_format(p_closing, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent=FIRST_LINE_INDENT, space_before=Pt(6), space_after=Pt(6))
    add_run_with_format(p_closing, p_closing.text, size=FONT_SIZE_DEFAULT, italic=True)


    # 7. Chữ ký
    format_signature_block(document, data)

    # 8. Nơi nhận (Tùy chọn)
    if data.get('recipients'):
         format_recipient_list(document, data)

    print("Định dạng Giấy mời hoàn tất.")
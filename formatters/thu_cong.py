# formatters/thu_cong.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Thư công không dùng common elements chuẩn NĐ30
    pass
except ImportError:
    pass
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_NAME

def format(document, data):
    print("Bắt đầu định dạng Thư công...")
    subject = data.get("subject", "Về việc trao đổi thông tin") # Tiêu đề thư
    body = data.get("body", "Nội dung thư công...")
    # Thông tin người gửi
    sender_info = data.get("sender_info", {"name": "[Tên người gửi]", "title": "[Chức vụ]", "org": "[Tên cơ quan/công ty]", "address": "[Địa chỉ]", "phone": "[Điện thoại]"})
    # Thông tin người nhận
    recipient_info = data.get("recipient_info", {"name": "[Tên người nhận]", "title": "[Chức vụ]", "org": "[Tên cơ quan/công ty]", "address": "[Địa chỉ]"})
    salutation = data.get("salutation", f"Kính gửi Ông/Bà {recipient_info.get('name', '')}:")
    closing = data.get("closing", "Trân trọng,")

    # 1. Header thư (Tên CQ gửi, địa chỉ, ngày tháng - căn phải)
    p_org_addr = document.add_paragraph()
    set_paragraph_format(p_org_addr, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=Pt(0), line_spacing=1.0)
    add_run_with_format(p_org_addr, sender_info.get('org', '').upper(), bold=True)
    p_org_addr.add_run("\n")
    add_run_with_format(p_org_addr, sender_info.get('address', ''))
    p_org_addr.add_run("\n")
    add_run_with_format(p_org_addr, f"{data.get('issuing_location', 'Hà Nội')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}", italic=True)

    # 2. Thông tin người nhận (Căn trái)
    p_rec_name = document.add_paragraph()
    set_paragraph_format(p_rec_name, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(12), space_after=Pt(0), line_spacing=1.0)
    add_run_with_format(p_rec_name, recipient_info.get('name', ''))
    if recipient_info.get('title'):
         p_rec_title = document.add_paragraph()
         set_paragraph_format(p_rec_title, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
         add_run_with_format(p_rec_title, recipient_info.get('title', ''))
    if recipient_info.get('org'):
         p_rec_org = document.add_paragraph()
         set_paragraph_format(p_rec_org, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
         add_run_with_format(p_rec_org, recipient_info.get('org', ''))
    if recipient_info.get('address'):
        p_rec_addr = document.add_paragraph()
        set_paragraph_format(p_rec_addr, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(6), line_spacing=1.0)
        add_run_with_format(p_rec_addr, recipient_info.get('address', ''))


    # 3. Chủ đề thư (V/v:)
    p_subject = document.add_paragraph()
    set_paragraph_format(p_subject, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(6), space_after=Pt(6))
    add_run_with_format(p_subject, f"V/v: {subject}", bold=True)

    # 4. Lời chào
    p_salutation = document.add_paragraph(salutation)
    set_paragraph_format(p_salutation, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(6))

    # 5. Nội dung thư
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    # 6. Lời kết
    p_closing = document.add_paragraph(closing)
    set_paragraph_format(p_closing, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(12), space_after=Pt(60)) # Chừa chỗ ký

    # 7. Chữ ký người gửi (Căn trái hoặc giữa dưới lời kết)
    p_signer_name = document.add_paragraph(sender_info.get('name', ''))
    set_paragraph_format(p_signer_name, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_signer_name, p_signer_name.text, bold=True)
    if sender_info.get('title'):
         p_signer_title = document.add_paragraph(sender_info.get('title', ''))
         set_paragraph_format(p_signer_title, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(0))
         add_run_with_format(p_signer_title, p_signer_title.text)


    print("Định dạng Thư công hoàn tất.")
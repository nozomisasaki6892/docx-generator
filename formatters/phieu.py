# formatters/phieu.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Phiếu thường không có header/footer/nơi nhận chuẩn
    pass
except ImportError:
    pass
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Phiếu...")
    # Xác định loại phiếu từ title hoặc data
    phieu_type = "PHIẾU"
    title = data.get("title", "Phiếu công tác")
    if title.upper().startswith("PHIẾU GỬI"): phieu_type = "PHIẾU GỬI"
    elif title.upper().startswith("PHIẾU CHUYỂN"): phieu_type = "PHIẾU CHUYỂN"
    elif title.upper().startswith("PHIẾU BÁO"): phieu_type = "PHIẾU BÁO"

    issuing_org = data.get("issuing_org", "TÊN ĐƠN VỊ").upper()
    issuing_dept = data.get("issuing_dept", None) # Bộ phận gửi

    # 1. Header đơn giản (Tên đơn vị, bộ phận)
    p_org = document.add_paragraph(issuing_org)
    set_paragraph_format(p_org, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_org, issuing_org, size=FONT_SIZE_HEADER, bold=True)
    if issuing_dept:
        p_dept = document.add_paragraph(issuing_dept.upper())
        set_paragraph_format(p_dept, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_dept, issuing_dept, size=Pt(11), bold=True)
    p_line = document.add_paragraph("-------***-------")
    set_paragraph_format(p_line, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 2. Tên loại Phiếu
    p_tenloai = document.add_paragraph(phieu_type)
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(12))
    add_run_with_format(p_tenloai, phieu_type, size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Thông tin cơ bản (Kính gửi, Từ, Ngày, V/v) - Có thể dùng bảng hoặc paragraph
    p_to = document.add_paragraph()
    set_paragraph_format(p_to, space_after=Pt(2))
    add_run_with_format(p_to, "Kính gửi:", bold=True)
    add_run_with_format(p_to, f" {data.get('phieu_to', '...................................................')}") # Chừa chỗ điền tay

    p_from = document.add_paragraph()
    set_paragraph_format(p_from, space_before=Pt(2), space_after=Pt(2))
    add_run_with_format(p_from, "Từ:", bold=True)
    add_run_with_format(p_from, f" {data.get('phieu_from', '...................................................')}")

    p_date_num = document.add_paragraph() # Ngày tháng và số phiếu (nếu có)
    set_paragraph_format(p_date_num, space_before=Pt(2), space_after=Pt(2))
    add_run_with_format(p_date_num, "Ngày:", bold=True)
    add_run_with_format(p_date_num, f" {time.strftime('%d/%m/%Y')}")
    if data.get('phieu_number'):
        add_run_with_format(p_date_num, f"      Số: {data.get('phieu_number')}")

    p_subject = document.add_paragraph()
    set_paragraph_format(p_subject, space_before=Pt(2), space_after=Pt(6))
    add_run_with_format(p_subject, "Về việc:", bold=True)
    add_run_with_format(p_subject, f" {data.get('phieu_subject', '...................................................')}")

    # 4. Nội dung phiếu (nếu có)
    body = data.get("body", None)
    if body:
        body_lines = body.split('\n')
        for line in body_lines:
            stripped_line = line.strip()
            if stripped_line:
                p = document.add_paragraph()
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
                add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    # 5. File đính kèm (nếu có)
    attachments = data.get("attachments", None)
    if attachments:
        p_attach_label = document.add_paragraph()
        set_paragraph_format(p_attach_label, space_before=Pt(6), space_after=Pt(0))
        add_run_with_format(p_attach_label, "Tài liệu kèm theo:", bold=True, italic=True)
        if isinstance(attachments, list):
            for attachment in attachments:
                p_att = document.add_paragraph()
                set_paragraph_format(p_att, left_indent=Cm(0.5), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
                add_run_with_format(p_att, f"- {attachment}", size=FONT_SIZE_DEFAULT, italic=True)
        else:
             p_att = document.add_paragraph()
             set_paragraph_format(p_att, left_indent=Cm(0.5), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
             add_run_with_format(p_att, f"- {attachments}", size=FONT_SIZE_DEFAULT, italic=True)


    # 6. Chữ ký người gửi/lập phiếu
    p_signer_title = document.add_paragraph()
    # Căn phải
    set_paragraph_format(p_signer_title, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(12), space_after=Pt(60)) # Chừa nhiều chỗ ký
    add_run_with_format(p_signer_title, data.get("signer_title", "NGƯỜI LẬP PHIẾU").upper(), bold=True)

    p_signer_name = document.add_paragraph()
    set_paragraph_format(p_signer_name, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(0))
    add_run_with_format(p_signer_name, data.get("signer_name", "[Ký, ghi rõ họ tên]"), bold=True)


    print("Định dạng Phiếu hoàn tất.")
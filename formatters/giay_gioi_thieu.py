# formatters/giay_gioi_thieu.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block
except ImportError:
    from common_elements import format_basic_header, format_signature_block
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Giấy giới thiệu...")
    # Thông tin cần thiết từ data
    introduced_person = data.get("introduced_person", {"name": "[Họ tên]", "title": "[Chức vụ]"})
    recipient_org = data.get("recipient_org", "[Tên cơ quan, đơn vị đến công tác]")
    purpose = data.get("purpose", "[Nội dung công tác]")
    valid_until = data.get("valid_until", "__/__/____") # ngày/tháng/năm

    # 1. Header chuẩn của cơ quan giới thiệu
    format_basic_header(document, data, "GiayGioiThieu")

    # 2. Tên loại
    p_tenloai = document.add_paragraph("GIẤY GIỚI THIỆU")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_tenloai, "GIẤY GIỚI THIỆU", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Kính gửi
    p_kg = document.add_paragraph(f"Kính gửi: {recipient_org}")
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12)) # Kính gửi căn giữa
    add_run_with_format(p_kg, p_kg.text, size=FONT_SIZE_DEFAULT, bold=True)

    # 4. Nội dung giới thiệu
    p_intro = document.add_paragraph()
    set_paragraph_format(p_intro, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_intro, "Trân trọng giới thiệu:")

    p_name = document.add_paragraph()
    set_paragraph_format(p_name, left_indent=Cm(1.0), space_after=Pt(0))
    add_run_with_format(p_name, f"Ông/Bà: {introduced_person['name']}", bold=True)

    p_title = document.add_paragraph()
    set_paragraph_format(p_title, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_title, f"Chức vụ: {introduced_person['title']}")

    p_to_org = document.add_paragraph()
    set_paragraph_format(p_to_org, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_to_org, f"Được cử đến công tác tại: {recipient_org}")

    p_purpose = document.add_paragraph()
    set_paragraph_format(p_purpose, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(6))
    add_run_with_format(p_purpose, f"Về việc: {purpose}")

    # 5. Đề nghị giúp đỡ
    p_request = document.add_paragraph()
    set_paragraph_format(p_request, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_request, f"Đề nghị Quý Cơ quan tạo điều kiện để Ông/Bà {introduced_person['name']} hoàn thành nhiệm vụ.")

    # 6. Hiệu lực
    p_validity = document.add_paragraph()
    set_paragraph_format(p_validity, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(12))
    add_run_with_format(p_validity, f"Giấy giới thiệu có giá trị đến hết ngày {valid_until}.")

    # 7. Chữ ký người đứng đầu cơ quan
    # Cần lấy đúng chức vụ từ data (signer_title)
    if not data.get('signer_title'): data['signer_title'] = "THỦ TRƯỞNG CƠ QUAN"
    format_signature_block(document, data)

    # Giấy giới thiệu không có nơi nhận ở cuối

    print("Định dạng Giấy giới thiệu hoàn tất.")
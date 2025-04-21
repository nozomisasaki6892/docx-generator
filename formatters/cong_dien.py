# formatters/cong_dien.py
import re
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Công điện có thể không dùng header/footer/nơi nhận chuẩn
    from .common_elements import format_signature_block
except ImportError:
    from common_elements import format_signature_block
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_NAME

def format(document, data):
    print("Bắt đầu định dạng Công điện...")
    title = data.get("title", "Công điện khẩn về việc ABC")
    body = data.get("body", "Nội dung công điện...")
    issuing_org = data.get("issuing_org", "CƠ QUAN BAN HÀNH").upper()
    recipients_cd = data.get("recipients_cd", "Điện gửi: [Danh sách nơi nhận điện]") # Nơi nhận đặc thù của CĐ
    urgency = data.get("urgency", "KHẨN").upper() # Hỏa tốc, Thượng khẩn, Khẩn

    # 1. Mức độ khẩn (Góc trên phải) - Dùng table
    urgency_table = document.add_table(rows=1, cols=2)
    urgency_table.autofit = False
    urgency_table.columns[0].width = Inches(4.0)
    urgency_table.columns[1].width = Inches(2.0)

    cell_left_blank = urgency_table.cell(0, 0)
    cell_left_blank._element.clear_content()
    cell_left_blank.add_paragraph("") # Ô trái trống

    cell_urgency = urgency_table.cell(0, 1)
    cell_urgency._element.clear_content()
    p_urgency = cell_urgency.add_paragraph(urgency)
    set_paragraph_format(p_urgency, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_urgency, urgency, size=FONT_SIZE_DEFAULT, bold=True)

    # 2. Tên loại CÔNG ĐIỆN
    p_tenloai = document.add_paragraph("CÔNG ĐIỆN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_tenloai, "CÔNG ĐIỆN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Tên cơ quan ban hành điện
    p_issuer = document.add_paragraph(issuing_org)
    set_paragraph_format(p_issuer, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_issuer, issuing_org, size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # 4. Nơi nhận điện
    p_recipients = document.add_paragraph(recipients_cd)
    # Thường căn trái và in đậm
    set_paragraph_format(p_recipients, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(12))
    add_run_with_format(p_recipients, recipients_cd, size=FONT_SIZE_DEFAULT, bold=True)

    # 5. Nội dung điện (Thường rất ngắn gọn, trực tiếp)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    # 6. Chữ ký (Căn phải, có thể có thời gian gửi cụ thể)
    current_time_str = time.strftime("%H giờ %M, ngày %d tháng %m năm %Y")
    p_time = document.add_paragraph(current_time_str)
    set_paragraph_format(p_time, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_time, current_time_str, size=Pt(11), italic=True)

    format_signature_block(document, data)

    # Công điện thường không có phần "Nơi nhận:" riêng ở cuối

    print("Định dạng Công điện hoàn tất.")
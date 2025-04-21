# formatters/thong_bao_nt.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Dùng header, signature, recipient chuẩn của trường/đơn vị
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Thông báo nhà trường...")
    title = data.get("title", "Thông báo về việc [Nội dung thông báo]")
    body = data.get("body", "Nội dung thông báo...")
    issuing_org = data.get("issuing_org", "TÊN TRƯỜNG/KHOA/PHÒNG BAN").upper()

    # 1. Header của trường/đơn vị
    data['issuing_org'] = issuing_org
    format_basic_header(document, data, "ThongBaoNT") # Dùng thể thức như Thông báo thường

    # 2. Tên loại
    p_tenloai = document.add_paragraph("THÔNG BÁO")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "THÔNG BÁO", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Trích yếu
    tb_title = title.replace("Thông báo", "").strip()
    p_title = document.add_paragraph(f"Về việc {tb_title}")
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, f"Về việc {tb_title}", size=Pt(14), bold=True)

    # 4. Nội dung thông báo
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            # Kiểm tra xem có phải mục đánh số/gạch đầu dòng không
            is_list_item = re.match(r'^(\d+\.|[a-z]\)|-)\s+', stripped_line)
            align = WD_ALIGN_PARAGRAPH.JUSTIFY
            left_indent = Cm(0.5) if is_list_item else Cm(0)
            first_indent = Cm(0) if is_list_item else FIRST_LINE_INDENT

            set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    # 5. Chữ ký (Lãnh đạo đơn vị thông báo)
    # Cần xác định chức danh ký phù hợp từ data
    if not data.get('signer_title'): data['signer_title'] = "TRƯỞNG KHOA/PHÒNG/BAN" # Ví dụ
    format_signature_block(document, data)

    # 6. Nơi nhận (Sinh viên, cán bộ, các đơn vị liên quan)
    if not data.get('recipients'): data['recipients'] = ["- Như trên;", "- Lưu: VT, [Đơn vị ban hành]."] # Ví dụ
    format_recipient_list(document, data)

    print("Định dạng Thông báo nhà trường hoàn tất.")
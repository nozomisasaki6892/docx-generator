# formatters/thong_bao_ts.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Thông báo TS dùng header và signature chuẩn của trường
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Thông báo tuyển sinh...")
    title = data.get("title", "Thông báo tuyển sinh Đại học/Cao đẳng/...")
    body = data.get("body", "Nội dung thông báo tuyển sinh...")
    # Tên trường/đơn vị tuyển sinh
    issuing_org = data.get("issuing_org", "TÊN TRƯỜNG/ĐƠN VỊ").upper()
    # Đơn vị cấp trên (nếu có, vd: Bộ Giáo dục và Đào tạo)
    issuing_org_parent = data.get("issuing_org_parent", None)

    # 1. Header của trường (cần truyền tên trường vào data)
    data['issuing_org'] = issuing_org
    if issuing_org_parent: data['issuing_org_parent'] = issuing_org_parent
    format_basic_header(document, data, "ThongBaoTS") # Dùng thể thức như Thông báo thường

    # 2. Tên loại
    p_tenloai = document.add_paragraph("THÔNG BÁO")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "THÔNG BÁO", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Trích yếu (Tiêu đề của thông báo tuyển sinh)
    ts_title = title.replace("Thông báo", "").strip()
    p_title = document.add_paragraph(ts_title)
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, ts_title, size=Pt(14), bold=True)

    # 4. Nội dung thông báo (Đối tượng, chỉ tiêu, ngành, hồ sơ, thời gian, địa điểm, lệ phí, liên hệ...)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            # Nhận diện các mục lớn (I, II, 1, 2) để in đậm
            is_main_heading = re.match(r'^[IVXLCDM]+\.\s+', stripped_line) or \
                              re.match(r'^\d+\.\s+', stripped_line)
            align = WD_ALIGN_PARAGRAPH.JUSTIFY
            indent = Cm(0)
            first_indent = FIRST_LINE_INDENT
            if is_main_heading:
                align = WD_ALIGN_PARAGRAPH.LEFT
                first_indent = Cm(0)

            set_paragraph_format(p, alignment=align, left_indent=indent, first_line_indent=first_indent, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=bool(is_main_heading))

    # 5. Chữ ký (Thường là Hiệu trưởng hoặc Chủ tịch Hội đồng tuyển sinh)
    if not data.get('signer_title'): data['signer_title'] = "HIỆU TRƯỞNG" # Hoặc CHỦ TỊCH HĐTS
    format_signature_block(document, data)

    # 6. Nơi nhận (Thí sinh, các đơn vị liên quan...)
    if not data.get('recipients'): data['recipients'] = ["- Như trên;", "- Lưu: VT, Phòng Đào tạo."] # Ví dụ
    format_recipient_list(document, data)

    print("Định dạng Thông báo tuyển sinh hoàn tất.")
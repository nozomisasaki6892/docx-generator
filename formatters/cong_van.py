# formatters/cong_van.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Import các hàm tiện ích và thành phần chung
from utils import set_paragraph_format, set_run_format, add_run_with_format
# Quan trọng: Đảm bảo common_elements.py cùng cấp hoặc chỉnh đường dẫn import
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    # Fallback nếu chạy trực tiếp file này (ít gặp)
    from common_elements import format_basic_header, format_signature_block, format_recipient_list

from config import FONT_SIZE_DEFAULT, FONT_SIZE_SMALL, FIRST_LINE_INDENT

def format(document, data):
    """Hàm chính định dạng cho loại văn bản Công Văn (Mặc định)."""
    print("Bắt đầu định dạng Công văn...")
    title = data.get("title", "Về việc ABC") # Thường là trích yếu
    body = data.get("body", "Nội dung công văn...") # Nội dung đã qua AI làm sạch
    recipients_to = data.get("recipients_to", "Kính gửi: [Tên đơn vị/cá nhân]") # Nên truyền từ data

    # 1. Tạo Header chuẩn cho Công văn (CQBH căn trái)
    # Truyền doc_type="CongVan" để format_basic_header biết cách căn lề CQBH
    format_basic_header(document, data, "CongVan")

    # 2. Trích yếu nội dung (V/v) - Căn trái, dưới Số/KH
    # Cần lấy đúng nội dung trích yếu từ title
    trich_yeu_text = title
    if title.lower().startswith("về việc"):
        trich_yeu_text = title[len("về việc"):].strip()
    elif title.lower().startswith("v/v"):
         trich_yeu_text = title[len("v/v"):].strip()

    p_vv = document.add_paragraph()
    # Canh trái, size 12pt, không đậm, không nghiêng
    set_paragraph_format(p_vv, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(6))
    run_vv = add_run_with_format(p_vv, f"V/v: {trich_yeu_text}", size=Pt(12))

    # 3. Kính gửi
    p_kg = document.add_paragraph()
    # Canh trái nhưng thụt vào so với lề (thường bằng vị trí bắt đầu nội dung)
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=FIRST_LINE_INDENT, space_before=Pt(6), space_after=Pt(6))
    run_kg = add_run_with_format(p_kg, recipients_to, size=FONT_SIZE_DEFAULT, bold=True) # Kính gửi đậm

    # 4. Nội dung chính
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            # Căn đều, thụt lề dòng đầu
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    # 5. Chữ ký
    format_signature_block(document, data)

    # 6. Nơi nhận
    format_recipient_list(document, data)

    print("Định dạng Công văn hoàn tất.")
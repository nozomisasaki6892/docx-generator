# formatters/chi_thi.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    """Hàm chính định dạng cho loại văn bản Chỉ Thị."""
    print("Bắt đầu định dạng Chỉ thị...")
    title = data.get("title", "Chỉ thị về việc ABC")
    body = data.get("body", "Nội dung chỉ thị...") # Nội dung đã qua AI làm sạch

    # 1. Tạo Header chuẩn (CQBH căn trái)
    format_basic_header(document, data, "ChiThi")

    # 2. Tên loại
    p_tenloai = document.add_paragraph("CHỈ THỊ")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "CHỈ THỊ", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Trích yếu
    trich_yeu_text = title.replace("Chỉ thị", "").strip()
    p_trichyeu = document.add_paragraph(f"Về việc {trich_yeu_text}")
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_trichyeu, f"Về việc {trich_yeu_text}", size=Pt(14), bold=True)
    # Chỉ thị thường không có gạch dưới trích yếu

    # 4. Nội dung chính (Thường là các đoạn mô tả và các mục chỉ đạo đánh số)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            is_numbered_item = re.match(r'^\d+\.\s+', stripped_line)
            align = WD_ALIGN_PARAGRAPH.JUSTIFY
            # Mục đánh số không thụt lề, đoạn văn thường thụt lề dòng đầu
            indent = Cm(0) if is_numbered_item else FIRST_LINE_INDENT
            first_indent_val = Cm(0) if is_numbered_item else FIRST_LINE_INDENT

            set_paragraph_format(p, alignment=align, left_indent=(indent if is_numbered_item else Cm(0)), first_line_indent=first_indent_val, line_spacing=1.5, space_after=Pt(6))

            # In đậm phần số thứ tự "1.", "2."...
            if is_numbered_item:
                 match = re.match(r'^(\d+\.)(\s+.*)', stripped_line)
                 if match:
                     add_run_with_format(p, match.group(1), size=FONT_SIZE_DEFAULT, bold=True)
                     add_run_with_format(p, match.group(2), size=FONT_SIZE_DEFAULT)
                 else: # Nếu chỉ có số (ít gặp)
                     add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=True)
            else: # Đoạn văn thường
                add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)


    # 5. Chữ ký
    format_signature_block(document, data)

    # 6. Nơi nhận
    format_recipient_list(document, data)

    print("Định dạng Chỉ thị hoàn tất.")
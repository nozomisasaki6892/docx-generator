# formatters/ke_hoach.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Import các hàm tiện ích và thành phần chung
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Thử import từ thư mục hiện tại (nếu chạy trong package)
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    # Fallback nếu chạy file trực tiếp
    from common_elements import format_basic_header, format_signature_block, format_recipient_list

from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    """Hàm chính định dạng cho loại văn bản Kế Hoạch."""
    print("Bắt đầu định dạng Kế hoạch...")
    title = data.get("title", "Kế hoạch thực hiện công việc ABC")
    body = data.get("body", "Nội dung kế hoạch...") # Nội dung đã qua AI làm sạch

    # 1. Tạo Header chuẩn (CQBH căn trái)
    format_basic_header(document, data, "KeHoach")

    # 2. Tên loại
    p_tenloai = document.add_paragraph("KẾ HOẠCH")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "KẾ HOẠCH", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Trích yếu (Tiêu đề của Kế hoạch)
    # Thường không có "Về việc"
    trich_yeu_text = title.replace("Kế hoạch", "").strip()
    p_trichyeu = document.add_paragraph(trich_yeu_text)
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_trichyeu, trich_yeu_text, size=Pt(14), bold=True)
    # Kế hoạch thường không có gạch dưới trích yếu

    # 4. Nội dung chính (Thường có các mục La Mã, Ả Rập, chữ cái, gạch đầu dòng...)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()

            # Xác định cấp độ mục và định dạng tương ứng
            left_indent_val = Cm(0)
            first_indent_val = Cm(0)
            is_bold_run = False
            align = WD_ALIGN_PARAGRAPH.JUSTIFY # Mặc định căn đều

            # Mục La Mã (I., II., ...)
            if re.match(r'^[IVXLCDM]+\.\s+', stripped_line):
                is_bold_run = True
                align = WD_ALIGN_PARAGRAPH.CENTER # Mục La Mã thường căn giữa hoặc trái+đậm
                set_paragraph_format(p, alignment=align, left_indent=Cm(0), first_line_indent=Cm(0), space_before=Pt(6), space_after=Pt(6))
                add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=True)
            # Mục Ả Rập (1., 2., ...)
            elif re.match(r'^\d+\.\s+', stripped_line):
                 is_bold_run = True # In đậm cả mục
                 left_indent_val = Cm(0.5)
                 set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=left_indent_val, first_line_indent=Cm(0), space_after=Pt(6))
                 add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=True)
            # Mục chữ cái (a), b), ...)
            elif re.match(r'^[a-z]\)\s+', stripped_line):
                 left_indent_val = Cm(1.0)
                 set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent_val, first_line_indent=Cm(0), space_after=Pt(6))
                 add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)
            # Mục gạch đầu dòng (-)
            elif stripped_line.startswith('-'):
                 left_indent_val = Cm(1.5)
                 set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent_val, first_line_indent=Cm(0), space_after=Pt(6))
                 # Giữ nguyên dấu gạch đầu dòng
                 add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)
            # Đoạn văn bản thường
            else:
                 first_indent_val = FIRST_LINE_INDENT
                 set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0), first_line_indent=first_indent_val, line_spacing=1.5, space_after=Pt(6))
                 add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)


    # 5. Chữ ký
    format_signature_block(document, data)

    # 6. Nơi nhận
    format_recipient_list(document, data)

    print("Định dạng Kế hoạch hoàn tất.")
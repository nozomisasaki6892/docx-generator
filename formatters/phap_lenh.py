# formatters/phap_lenh.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
# Giả sử dùng header và signature của Luật nhưng CQBH và chức ký khác
from .luat import format_qppl_header, format_qppl_signature
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Pháp lệnh...")
    title = data.get("title", "Pháp lệnh về ABC")
    body = data.get("body", "Nội dung pháp lệnh...")
    ordinance_number = data.get("ordinance_number", "Pháp lệnh số: .../.../UBTVQH...")
    adoption_date_str = data.get("adoption_date", time.strftime("ngày %d tháng %m năm %Y"))

    # 1. Header (CQBH là UBTVQH)
    format_qppl_header(document, "ỦY BAN THƯỜNG VỤ QUỐC HỘI")

    # 2. Số hiệu Pháp lệnh
    p_num = document.add_paragraph(ordinance_number)
    set_paragraph_format(p_num, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_num, p_num.text, size=FONT_SIZE_DEFAULT)

    # 3. Tên Pháp lệnh
    pl_title = title.upper()
    p_tenloai = document.add_paragraph(pl_title)
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(12))
    add_run_with_format(p_tenloai, p_tenloai.text, size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Căn cứ ban hành
    preamble = data.get("preamble", "Căn cứ Hiến pháp nước Cộng hòa xã hội chủ nghĩa Việt Nam;\nCăn cứ Luật Tổ chức Quốc hội;\n[Căn cứ khác];\nỦy ban Thường vụ Quốc hội ban hành Pháp lệnh ...")
    preamble_lines = preamble.split('\n')
    for line in preamble_lines:
         p_pre = document.add_paragraph(line)
         set_paragraph_format(p_pre, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(0), line_spacing=1.5)
         add_run_with_format(p_pre, line, size=FONT_SIZE_DEFAULT, italic=True)
    document.add_paragraph() # Thêm khoảng trống

    # 5. Nội dung (Chương, Mục, Điều, Khoản, Điểm) - Tương tự Luật
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if not stripped_line: continue
        p = document.add_paragraph()
        # (Copy logic xử lý Chương, Mục, Điều, Khoản, Điểm từ formatters/luat.py)
        is_chuong = stripped_line.upper().startswith("CHƯƠNG")
        is_muc = re.match(r'^(MỤC\s+\d+)\.?\s+', stripped_line.upper())
        is_dieu = stripped_line.upper().startswith("ĐIỀU")
        is_khoan = re.match(r'^\d+\.\s+', stripped_line)
        is_diem = re.match(r'^[a-z]\)\s+', stripped_line)

        align = WD_ALIGN_PARAGRAPH.JUSTIFY
        left_indent = Cm(0)
        first_indent = FIRST_LINE_INDENT
        is_bold = False
        size = FONT_SIZE_DEFAULT
        space_before = Pt(0)
        space_after = Pt(6)

        if is_chuong:
            align = WD_ALIGN_PARAGRAPH.CENTER
            first_indent = Cm(0)
            is_bold = True
            space_before = Pt(12)
        elif is_muc:
            align = WD_ALIGN_PARAGRAPH.CENTER
            first_indent = Cm(0)
            is_bold = True
            space_before = Pt(6)
        elif is_dieu:
            align = WD_ALIGN_PARAGRAPH.LEFT
            first_indent = Cm(0)
            is_bold = True
            space_before = Pt(6)
        elif is_khoan:
            left_indent = Cm(0.5)
            first_indent = Cm(0)
        elif is_diem:
            left_indent = Cm(1.0)
            first_indent = Cm(0)

        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_before=space_before, space_after=space_after)
        add_run_with_format(p, stripped_line, size=size, bold=is_bold)


    # 6. Thông tin thông qua Pháp lệnh
    p_adoption = document.add_paragraph(f"Pháp lệnh này đã được Ủy ban Thường vụ Quốc hội nước Cộng hòa xã hội chủ nghĩa Việt Nam khóa ... thông qua ngày {adoption_date_str}.")
    set_paragraph_format(p_adoption, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_adoption, p_adoption.text, size=FONT_SIZE_DEFAULT, italic=True)

    # 7. Chữ ký (TM. UBTVQH, CHỦ TỊCH)
    format_qppl_signature(document, "TM. ỦY BAN THƯỜNG VỤ QUỐC HỘI\nCHỦ TỊCH", data.get("signer_name", "[Tên Chủ tịch QH]"))

    print("Định dạng Pháp lệnh hoàn tất.")
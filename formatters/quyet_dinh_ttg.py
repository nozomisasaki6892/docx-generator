# formatters/quyet_dinh_ttg.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
# Header TTg giống CP, Chữ ký khác
from .luat import format_qppl_header
from .nghi_dinh_qppl import format_cp_signature # Chữ ký TTg tương tự NĐ
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Quyết định TTg...")
    title = data.get("title", "Quyết định của Thủ tướng Chính phủ về ABC")
    body = data.get("body", "Nội dung quyết định...")
    decision_number = data.get("decision_number", "Quyết định số: .../20.../QĐ-TTg")
    issuing_date_str = data.get("issuing_date", time.strftime("ngày %d tháng %m năm %Y"))
    issuing_location = data.get("issuing_location", "Hà Nội")

    # 1. Header (THỦ TƯỚNG CHÍNH PHỦ)
    format_qppl_header(document, "THỦ TƯỚNG CHÍNH PHỦ")

    # 2. Số hiệu và Ngày tháng ban hành (căn giữa)
    p_num_date = document.add_paragraph(f"{decision_number}\n{issuing_location}, {issuing_date_str}")
    set_paragraph_format(p_num_date, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_num_date, p_num_date.text, size=FONT_SIZE_DEFAULT)

    # 3. Tên Quyết định
    qd_title = title.replace("của Thủ tướng Chính phủ", "").strip().upper()
    p_tenloai = document.add_paragraph(qd_title)
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(12))
    add_run_with_format(p_tenloai, p_tenloai.text, size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Cơ quan ban hành (THỦ TƯỚNG CHÍNH PHỦ)
    p_issuer = document.add_paragraph("THỦ TƯỚNG CHÍNH PHỦ")
    set_paragraph_format(p_issuer, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_issuer, p_issuer.text, size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # 5. Căn cứ ban hành
    preamble = data.get("preamble", "Căn cứ Luật Tổ chức Chính phủ ngày ... tháng ... năm ...;\nCăn cứ [Luật/Pháp lệnh/Nghị định liên quan];\nXét đề nghị của [Bộ trưởng/Thủ trưởng cơ quan];\nThủ tướng Chính phủ ban hành Quyết định ...")
    preamble_lines = preamble.split('\n')
    for line in preamble_lines:
         p_pre = document.add_paragraph(line)
         set_paragraph_format(p_pre, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(0), line_spacing=1.5)
         add_run_with_format(p_pre, line, size=FONT_SIZE_DEFAULT, italic=True)
    document.add_paragraph()

    # 6. Nội dung (Thường chỉ có Điều, Khoản, Điểm)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if not stripped_line: continue
        p = document.add_paragraph()
        # (Copy logic xử lý Điều, Khoản, Điểm từ formatters/nghi_quyet_qh.py)
        is_dieu = stripped_line.upper().startswith("ĐIỀU")
        is_khoan = re.match(r'^\d+\.\s+', stripped_line)
        is_diem = re.match(r'^[a-z]\)\s+', stripped_line)

        align = WD_ALIGN_PARAGRAPH.JUSTIFY
        left_indent = Cm(0)
        first_indent = FIRST_LINE_INDENT
        is_bold = False

        if is_dieu:
            align = WD_ALIGN_PARAGRAPH.LEFT
            first_indent = Cm(0)
            is_bold = True
        elif is_khoan:
            left_indent = Cm(0.5)
            first_indent = Cm(0)
        elif is_diem:
            left_indent = Cm(1.0)
            first_indent = Cm(0)

        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_after=Pt(6))
        add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_bold)


    # 7. Chữ ký (THỦ TƯỚNG) - Chữ ký trực tiếp, không có TM.
    # Sử dụng lại hàm của NĐ CP nhưng bỏ TM.
    sig_paragraph = document.add_paragraph()
    set_paragraph_format(sig_paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(6), space_after=Pt(0), line_spacing=1.15)
    add_run_with_format(sig_paragraph, "THỦ TƯỚNG\n\n\n\n\n", size=FONT_SIZE_SIGNATURE, bold=True)
    add_run_with_format(sig_paragraph, data.get("signer_name", "[Tên Thủ tướng]"), size=FONT_SIZE_SIGNER_NAME, bold=True)


    # Quyết định TTg thường không có nơi nhận ở cuối

    print("Định dạng Quyết định TTg hoàn tất.")
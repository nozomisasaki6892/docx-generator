import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

try:
    from .luat import format_qppl_header
except ImportError:
    from formatters.luat import format_qppl_header

from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_SIGNATURE, FONT_SIZE_SIGNER_NAME, FONT_SIZE_PLACE_DATE

def format_cp_signature(document, signer_title, signer_name):
    sig_paragraph = document.add_paragraph()
    set_paragraph_format(sig_paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(6), space_after=Pt(0), line_spacing=1.15)
    add_run_with_format(sig_paragraph, "TM. CHÍNH PHỦ\n", size=FONT_SIZE_SIGNATURE, bold=True)
    add_run_with_format(sig_paragraph, signer_title.upper() + "\n\n\n\n\n", size=FONT_SIZE_SIGNATURE, bold=True)
    add_run_with_format(sig_paragraph, signer_name, size=FONT_SIZE_SIGNER_NAME, bold=True)

def format(document, data):
    print("Bắt đầu định dạng Nghị định...")
    title = data.get("title", "Nghị định").upper()
    body = data.get("body", "")
    decree_number = data.get("decree_number", "Số: .../20.../NĐ-CP")
    issuing_date_str = data.get("issuing_date", time.strftime("ngày %d tháng %m năm %Y"))
    issuing_location = data.get("issuing_location", "Hà Nội")
    issuer_name = "CHÍNH PHỦ"
    preamble = data.get("preamble", []) # Mặc định là list rỗng nếu không có

    format_qppl_header(document, issuer_name)

    p_num_date = document.add_paragraph()
    set_paragraph_format(p_num_date, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_num_date, f"{decree_number} ", size=FONT_SIZE_DEFAULT)
    add_run_with_format(p_num_date, f"{issuing_location}, {issuing_date_str}", size=FONT_SIZE_DEFAULT, italic=True)

    p_tenloai = document.add_paragraph()
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(12))
    add_run_with_format(p_tenloai, title, size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    p_issuer_body = document.add_paragraph()
    set_paragraph_format(p_issuer_body, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_issuer_body, issuer_name, size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # Xử lý phần căn cứ, đề nghị (preamble)
    if preamble:
        for line in preamble:
            stripped_line = line.strip()
            if stripped_line:
                p = document.add_paragraph()
                set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(0))
                add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, italic=True)
        document.add_paragraph() # Thêm khoảng trống sau phần căn cứ

    # Xử lý phần nội dung chính (body)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if not stripped_line:
            continue

        p = document.add_paragraph()
        is_chuong = stripped_line.upper().startswith("CHƯƠNG")
        is_muc = re.match(r'^(MỤC\s+\d+)\.?\s+', stripped_line.upper())
        is_dieu = stripped_line.upper().startswith("ĐIỀU")
        is_khoan = re.match(r'^\d+\.\s+', stripped_line)
        is_diem = re.match(r'^[a-z]\)\s+', stripped_line)

        align = WD_ALIGN_PARAGRAPH.JUSTIFY
        left_indent = Cm(0)
        first_indent = FIRST_LINE_INDENT
        is_bold = False
        is_italic = False
        size = FONT_SIZE_DEFAULT
        space_before = Pt(0)
        space_after = Pt(6)
        line_spacing = 1.5

        if is_chuong:
            align = WD_ALIGN_PARAGRAPH.CENTER
            is_bold = True
            space_before = Pt(12)
            space_after = Pt(6) # Giữ khoảng cách sau Chương
            size = Pt(13)
            first_indent = Cm(0) # Chương không thụt đầu dòng
            # Tách số và tên chương nếu có
            match = re.match(r'^(CHƯƠNG\s+[IVXLCDM]+)\s*-\s*(.*)', stripped_line.upper()) or \
                    re.match(r'^(CHƯƠNG\s+[IVXLCDM]+)\s+(.*)', stripped_line.upper())
            if match:
                 chuong_num = match.group(1)
                 chuong_title = match.group(2)
                 add_run_with_format(p, chuong_num, size=size, bold=is_bold)
                 p.add_run("\n") # Xuống dòng
                 add_run_with_format(p, chuong_title.upper(), size=size, bold=is_bold) # Tên chương viết hoa
            else:
                 add_run_with_format(p, stripped_line.upper(), size=size, bold=is_bold) # Nếu không tách được thì giữ nguyên

        elif is_muc:
            align = WD_ALIGN_PARAGRAPH.CENTER
            is_bold = True
            space_before = Pt(6)
            space_after = Pt(6) # Giữ khoảng cách sau Mục
            size = Pt(13)
            first_indent = Cm(0) # Mục không thụt đầu dòng
             # Tách số và tên mục
            match = re.match(r'^(MỤC\s+\d+)\.?\s+(.*)', stripped_line.upper())
            if match:
                 muc_num = match.group(1)
                 muc_title = match.group(2)
                 add_run_with_format(p, muc_num, size=size, bold=is_bold)
                 p.add_run("\n") # Xuống dòng
                 add_run_with_format(p, muc_title.upper(), size=size, bold=is_bold) # Tên mục viết hoa
            else:
                 add_run_with_format(p, stripped_line.upper(), size=size, bold=is_bold)

        elif is_dieu:
            align = WD_ALIGN_PARAGRAPH.LEFT # Điều căn trái
            is_bold = True
            space_before = Pt(6)
            space_after = Pt(3) # Giảm khoảng cách sau Điều
            size = Pt(13)
            first_indent = Cm(0) # Điều không thụt đầu dòng
             # Tách số và tên điều
            match = re.match(r'^(ĐIỀU\s+\d+)\.?\s+(.*)', stripped_line, re.IGNORECASE)
            if match:
                dieu_num_title = match.group(1) # Giữ nguyên case "Điều"
                dieu_content_start = match.group(2)
                run_num = add_run_with_format(p, f"{dieu_num_title}. ", size=size, bold=is_bold)
                run_content = add_run_with_format(p, dieu_content_start, size=size, bold=False) # Nội dung tên điều không đậm
            else:
                 add_run_with_format(p, stripped_line, size=size, bold=is_bold)

        elif is_khoan:
            left_indent = FIRST_LINE_INDENT # Khoản thụt lề bằng first_line_indent
            first_indent = Cm(0) # Khoản không thụt dòng đầu tiên
            space_before = Pt(3) # Giảm khoảng cách trước khoản
            space_after = Pt(3)  # Giảm khoảng cách sau khoản
            match = re.match(r'^(\d+\.)(\s+)(.*)', stripped_line)
            if match:
                khoan_num = match.group(1)
                whitespace = match.group(2)
                khoan_content = match.group(3)
                run_num = add_run_with_format(p, khoan_num, size=size)
                # run_space = p.add_run(whitespace) # Giữ nguyên khoảng trắng
                # run_space.font.size = size
                run_content = add_run_with_format(p, " " + khoan_content, size=size) # Thêm 1 space sau số khoản
            else:
                 add_run_with_format(p, stripped_line, size=size)

        elif is_diem:
            left_indent = FIRST_LINE_INDENT + Cm(0.5) # Điểm thụt lề thêm 0.5cm so với Khoản
            first_indent = Cm(0) # Điểm không thụt dòng đầu tiên
            space_before = Pt(3) # Giảm khoảng cách trước điểm
            space_after = Pt(3)  # Giảm khoảng cách sau điểm
            match = re.match(r'^([a-z]\))(\s+)(.*)', stripped_line)
            if match:
                diem_marker = match.group(1)
                whitespace = match.group(2)
                diem_content = match.group(3)
                run_marker = add_run_with_format(p, diem_marker, size=size)
                # run_space = p.add_run(whitespace) # Giữ nguyên khoảng trắng
                # run_space.font.size = size
                run_content = add_run_with_format(p, " " + diem_content, size=size) # Thêm 1 space sau ký tự điểm
            else:
                add_run_with_format(p, stripped_line, size=size)

        else: # Đoạn văn bản thông thường
            if not (is_chuong or is_muc or is_dieu): # Chỉ áp dụng cho đoạn thường, không phải tiêu đề
                 add_run_with_format(p, stripped_line, size=size)

        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=line_spacing, space_before=space_before, space_after=space_after)

    signer_title = data.get("signer_title", "THỦ TƯỚNG")
    signer_name = data.get("signer_name", "[Tên Thủ tướng]")
    format_cp_signature(document, signer_title, signer_name)

    print("Định dạng Nghị định hoàn tất.")

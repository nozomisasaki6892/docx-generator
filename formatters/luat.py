# formatters/luat.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format_qppl_header(document, issuer_name):
    # Header VBQPPL thường chỉ có QH/TN căn giữa
    p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_qh, p_qh.text, size=FONT_SIZE_HEADER, bold=True)
    p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_tn, p_tn.text, size=Pt(14), bold=True) # Tiêu ngữ to hơn
    p_line_tn = document.add_paragraph("-" * 20)
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    # Tên cơ quan ban hành dưới QH/TN
    p_issuer = document.add_paragraph(issuer_name.upper())
    set_paragraph_format(p_issuer, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_issuer, p_issuer.text, size=FONT_SIZE_HEADER, bold=False) # Tên CQ không đậm


def format_qppl_signature(document, signer_title, signer_name):
     # Chữ ký VBQPPL thường căn giữa hoặc phải tùy loại
     sig_paragraph = document.add_paragraph()
     set_paragraph_format(sig_paragraph, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(0), line_spacing=1.15)
     add_run_with_format(sig_paragraph, signer_title.upper() + "\n\n\n\n\n", size=Pt(14), bold=True) # Chức vụ đậm, to
     add_run_with_format(sig_paragraph, signer_name, size=Pt(14), bold=True) # Tên đậm, to


def format(document, data):
    print("Bắt đầu định dạng Luật/Bộ luật...")
    title = data.get("title", "LUẬT ABC") # VD: LUẬT AN NINH MẠNG
    body = data.get("body", "Nội dung luật...")
    law_number = data.get("law_number", "Luật số: .../.../QH...") # VD: Luật số: 24/2018/QH14
    adoption_date_str = data.get("adoption_date", time.strftime("ngày %d tháng %m năm %Y")) # Ngày QH thông qua

    # 1. Header (Tên cơ quan ban hành là QUỐC HỘI)
    format_qppl_header(document, "QUỐC HỘI")

    # 2. Số hiệu Luật
    p_num = document.add_paragraph(law_number)
    set_paragraph_format(p_num, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_num, p_num.text, size=FONT_SIZE_DEFAULT)

    # 3. Tên Luật
    p_tenluat = document.add_paragraph(title.upper())
    set_paragraph_format(p_tenluat, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(12))
    add_run_with_format(p_tenluat, p_tenluat.text, size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Lời nói đầu / Căn cứ ban hành
    # Cần tách riêng hoặc nhận từ data['preamble']
    preamble = data.get("preamble", f"Căn cứ Hiến pháp nước Cộng hòa xã hội chủ nghĩa Việt Nam;\nQuốc hội ban hành {title.split(' ', 1)[-1]}.") # Mẫu preamble
    preamble_lines = preamble.split('\n')
    for line in preamble_lines:
         p_pre = document.add_paragraph(line)
         set_paragraph_format(p_pre, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
         add_run_with_format(p_pre, line, size=FONT_SIZE_DEFAULT, italic=True) # Căn cứ in nghiêng

    # 5. Nội dung (Chương, Mục, Điều, Khoản, Điểm)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if not stripped_line: continue
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
            align = WD_ALIGN_PARAGRAPH.LEFT # Điều thường căn trái
            first_indent = Cm(0)
            is_bold = True
            space_before = Pt(6)
        elif is_khoan:
            left_indent = Cm(0.5) # Thụt lề khoản
            first_indent = Cm(0)
        elif is_diem:
            left_indent = Cm(1.0) # Thụt lề điểm
            first_indent = Cm(0)

        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_before=space_before, space_after=space_after)
        add_run_with_format(p, stripped_line, size=size, bold=is_bold)

    # 6. Thông tin thông qua Luật
    p_adoption = document.add_paragraph(f"{title} này đã được Quốc hội nước Cộng hòa xã hội chủ nghĩa Việt Nam khóa ... kỳ họp thứ ... thông qua ngày {adoption_date_str}.")
    set_paragraph_format(p_adoption, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_adoption, p_adoption.text, size=FONT_SIZE_DEFAULT, italic=True)

    # 7. Chữ ký Chủ tịch Quốc hội
    format_qppl_signature(document, "CHỦ TỊCH QUỐC HỘI", data.get("signer_name", "[Tên Chủ tịch QH]"))

    print("Định dạng Luật hoàn tất.")
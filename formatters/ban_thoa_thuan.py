# formatters/ban_thoa_thuan.py
import re
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_signature_block # Tạm dùng, cần tùy chỉnh
except ImportError:
    from common_elements import format_signature_block
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Bản thỏa thuận...")
    title = data.get("title", "Bản thỏa thuận về vấn đề ABC")
    body = data.get("body", "Nội dung thỏa thuận...")
    party_a_info = data.get("party_a", {"name": "BÊN THỨ NHẤT (BÊN A)", "details": ["Thông tin Bên A..."]})
    party_b_info = data.get("party_b", {"name": "BÊN THỨ HAI (BÊN B)", "details": ["Thông tin Bên B..."]})

    # 1. Header (Tương tự MOU, có thể tùy chọn)
    add_qh_tn = data.get("add_qh_tn_agreement", True)
    if add_qh_tn:
        p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_qh, p_qh.text, size=FONT_SIZE_HEADER, bold=True)
        p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
        set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True)
        p_line_tn = document.add_paragraph("-" * 20)
        set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 2. Tên loại
    p_tenloai = document.add_paragraph("BẢN THỎA THUẬN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "BẢN THỎA THUẬN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)
    # (Có thể có tên tiếng Anh: AGREEMENT)

    # 3. Tiêu đề/Trích yếu của thỏa thuận
    agreement_title = title.replace("Bản thỏa thuận", "").strip()
    p_title = document.add_paragraph(agreement_title)
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, agreement_title, size=Pt(14), bold=True)

    # 4. Căn cứ ký thỏa thuận (nếu có)
    # Cần logic tách căn cứ từ body hoặc data riêng

    # 5. Thông tin các bên
    p_party_a_name = document.add_paragraph()
    set_paragraph_format(p_party_a_name, space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_party_a_name, party_a_info['name'], bold=True)
    for detail in party_a_info['details']:
         p_detail = document.add_paragraph()
         set_paragraph_format(p_detail, left_indent=Cm(0.5), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
         add_run_with_format(p_detail, detail, size=FONT_SIZE_DEFAULT)

    p_party_b_name = document.add_paragraph()
    set_paragraph_format(p_party_b_name, space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_party_b_name, party_b_info['name'], bold=True)
    for detail in party_b_info['details']:
         p_detail = document.add_paragraph()
         set_paragraph_format(p_detail, left_indent=Cm(0.5), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
         add_run_with_format(p_detail, detail, size=FONT_SIZE_DEFAULT)

    document.add_paragraph("Sau khi bàn bạc, hai bên thống nhất ký kết Bản thỏa thuận này với các điều khoản sau:", space_before=Pt(12))

    # 6. Nội dung thỏa thuận (Các Điều khoản)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            is_dieu = stripped_line.upper().startswith("ĐIỀU")
            is_khoan = re.match(r'^\d+\.\s+', stripped_line)
            align = WD_ALIGN_PARAGRAPH.LEFT if is_dieu else WD_ALIGN_PARAGRAPH.JUSTIFY
            indent = Cm(0.5) if is_khoan else Cm(0)
            first_indent = Cm(0) if (is_dieu or is_khoan) else FIRST_LINE_INDENT
            set_paragraph_format(p, alignment=align, left_indent=indent, first_line_indent=first_indent, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_dieu)

    # 7. Điều khoản chung (Hiệu lực, số bản, giải quyết tranh chấp...)
    document.add_paragraph("Thỏa thuận này được lập thành ... (...) bản có giá trị như nhau, mỗi bên giữ ... bản và có hiệu lực từ ngày ký.", space_before=Pt(12))

    # 8. Chữ ký các bên (Tương tự MOU)
    sig_table = document.add_table(rows=1, cols=2)
    sig_table.autofit = False
    sig_table.columns[0].width = Inches(3.0)
    sig_table.columns[1].width = Inches(3.0)

    # Chữ ký Bên A
    cell_a = sig_table.cell(0, 0)
    cell_a._element.clear_content()
    p_a_title = cell_a.add_paragraph(party_a_info['name']) # Hoặc "ĐẠI DIỆN BÊN A"
    set_paragraph_format(p_a_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_a_title, p_a_title.text, bold=True)

    # Chữ ký Bên B
    cell_b = sig_table.cell(0, 1)
    cell_b._element.clear_content()
    p_b_title = cell_b.add_paragraph(party_b_info['name']) # Hoặc "ĐẠI DIỆN BÊN B"
    set_paragraph_format(p_b_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_b_title, p_b_title.text, bold=True)

    print("Định dạng Bản thỏa thuận hoàn tất.")
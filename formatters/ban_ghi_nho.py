# formatters/ban_ghi_nho.py
import re
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Có thể dùng signature block nhưng layout chữ ký khác
    from .common_elements import format_signature_block # Tạm dùng, cần tùy chỉnh
except ImportError:
    from common_elements import format_signature_block
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Bản ghi nhớ (MOU)...")
    title = data.get("title", "Bản ghi nhớ hợp tác")
    body = data.get("body", "Nội dung ghi nhớ...")
    # Thông tin các bên ký kết (cần cấu trúc từ data)
    party_a_info = data.get("party_a", {"name": "BÊN A", "details": ["Đại diện bởi:...", "Chức vụ:...", "Địa chỉ:...", "Điện thoại:..."]})
    party_b_info = data.get("party_b", {"name": "BÊN B", "details": ["Đại diện bởi:...", "Chức vụ:...", "Địa chỉ:...", "Điện thoại:..."]})

    # 1. Header (Thường không theo NĐ30, có thể bỏ trống hoặc thêm logo)
    # Tạm bỏ qua header chuẩn

    # 2. Quốc hiệu/Tiêu ngữ (Có thể có hoặc không)
    add_qh_tn = data.get("add_qh_tn_mou", True) # Cho phép tùy chọn
    if add_qh_tn:
        p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_qh, p_qh.text, size=FONT_SIZE_HEADER, bold=True)
        p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
        set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True)
        p_line_tn = document.add_paragraph("-" * 20)
        set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 3. Tên loại
    p_tenloai = document.add_paragraph("BẢN GHI NHỚ")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "BẢN GHI NHỚ", size=FONT_SIZE_TITLE, bold=True, uppercase=True)
    # (Có thể thêm tên tiếng Anh: MEMORANDUM OF UNDERSTANDING)
    p_en_title = document.add_paragraph("MEMORANDUM OF UNDERSTANDING")
    set_paragraph_format(p_en_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_en_title, p_en_title.text, size=FONT_SIZE_TITLE, bold=True, uppercase=True)


    # 4. Ngày tháng ký (có thể đặt ở đây hoặc cuối)
    p_date_place = document.add_paragraph(f"{data.get('issuing_location', 'Hà Nội')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_date_place, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_date_place, p_date_place.text, size=FONT_SIZE_DEFAULT, italic=True)

    # 5. Thông tin các bên
    document.add_paragraph("Hôm nay, ngày ... tháng ... năm ..., tại ..., chúng tôi gồm:") # Phần dẫn nhập
    # Bên A
    p_party_a_name = document.add_paragraph()
    set_paragraph_format(p_party_a_name, left_indent=Cm(0.5), space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_party_a_name, party_a_info['name'], bold=True)
    for detail in party_a_info['details']:
        p_detail = document.add_paragraph()
        set_paragraph_format(p_detail, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_detail, detail, size=FONT_SIZE_DEFAULT)
    # Bên B
    p_party_b_name = document.add_paragraph()
    set_paragraph_format(p_party_b_name, left_indent=Cm(0.5), space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_party_b_name, party_b_info['name'], bold=True)
    for detail in party_b_info['details']:
        p_detail = document.add_paragraph()
        set_paragraph_format(p_detail, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_detail, detail, size=FONT_SIZE_DEFAULT)

    document.add_paragraph("Hai bên cùng thống nhất ký kết Bản ghi nhớ này với các nội dung sau:", space_before=Pt(12))

    # 6. Nội dung ghi nhớ (Thường là các Điều khoản)
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

    # 7. Hiệu lực, số bản
    document.add_paragraph("Bản ghi nhớ này có hiệu lực kể từ ngày ký và được lập thành ... (...) bản có giá trị pháp lý như nhau, mỗi bên giữ ... bản.", space_before=Pt(12))

    # 8. Chữ ký các bên (Dùng table 2 cột)
    sig_table = document.add_table(rows=1, cols=2)
    sig_table.autofit = False
    sig_table.columns[0].width = Inches(3.0)
    sig_table.columns[1].width = Inches(3.0)

    # Chữ ký Bên A
    cell_a = sig_table.cell(0, 0)
    cell_a._element.clear_content()
    p_a_title = cell_a.add_paragraph(party_a_info['name'])
    set_paragraph_format(p_a_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60)) # Chừa chỗ ký
    add_run_with_format(p_a_title, party_a_info['name'], bold=True)
    # Tên người đại diện A (nếu có)
    # p_a_signer = cell_a.add_paragraph(party_a_info['details'][0].split(':')[-1].strip()) # Ví dụ lấy tên từ dòng đại diện
    # set_paragraph_format(p_a_signer, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    # add_run_with_format(p_a_signer, p_a_signer.text, bold=True)


    # Chữ ký Bên B
    cell_b = sig_table.cell(0, 1)
    cell_b._element.clear_content()
    p_b_title = cell_b.add_paragraph(party_b_info['name'])
    set_paragraph_format(p_b_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60)) # Chừa chỗ ký
    add_run_with_format(p_b_title, party_b_info['name'], bold=True)
    # Tên người đại diện B (nếu có)
    # p_b_signer = cell_b.add_paragraph(party_b_info['details'][0].split(':')[-1].strip())
    # set_paragraph_format(p_b_signer, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    # add_run_with_format(p_b_signer, p_b_signer.text, bold=True)


    print("Định dạng Bản ghi nhớ hoàn tất.")
# formatters/hop_dong.py
import re
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    pass # Hợp đồng không dùng common elements của NĐ30
except ImportError:
    pass
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER, FONT_SIZE_SIGNATURE, FONT_SIZE_SIGNER_NAME

def format(document, data):
    print("Bắt đầu định dạng Hợp đồng...")
    contract_type = data.get("contract_type", "KINH TẾ") # VD: Kinh tế, Dịch vụ, Lao động,...
    title = data.get("title", f"Hợp đồng {contract_type}")
    contract_number = data.get("contract_number", "Số: ...... /HĐKT")
    body = data.get("body", "Nội dung hợp đồng...")
    party_a_info = data.get("party_a", {"name": "BÊN A:", "details": ["Tên công ty/cá nhân:...", "Địa chỉ:", "Mã số thuế:", "Đại diện bởi:", "Chức vụ:"]})
    party_b_info = data.get("party_b", {"name": "BÊN B:", "details": ["Tên công ty/cá nhân:...", "Địa chỉ:", "Mã số thuế:", "Đại diện bởi:", "Chức vụ:"]})

    # 1. Header (QH/TN và Tên công ty/Số HĐ) - Dùng table
    header_table = document.add_table(rows=1, cols=2)
    header_table.autofit = False
    header_table.columns[0].width = Inches(3.0)
    header_table.columns[1].width = Inches(3.0)

    # Cột trái: Tên Công ty A (Ví dụ) hoặc để trống
    cell_org = header_table.cell(0, 0)
    cell_org._element.clear_content()
    # p_org_a = cell_org.add_paragraph(party_a_info['details'][0].split(':')[-1].strip().upper())
    # set_paragraph_format(p_org_a, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    # add_run_with_format(p_org_a, p_org_a.text, size=FONT_SIZE_HEADER, bold=True)
    # p_org_a_line = cell_org.add_paragraph("-------")
    # set_paragraph_format(p_org_a_line, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))

    # Cột phải: QH/TN
    cell_qh_tn = header_table.cell(0, 1)
    cell_qh_tn._element.clear_content()
    p_qh = cell_qh_tn.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_qh, p_qh.text, size=FONT_SIZE_HEADER, bold=True)
    p_tn = cell_qh_tn.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True)
    p_line_tn = cell_qh_tn.add_paragraph("-" * 20)
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))

    # Số hợp đồng (dưới QH/TN)
    p_hd_num = cell_qh_tn.add_paragraph(contract_number)
    set_paragraph_format(p_hd_num, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6))
    add_run_with_format(p_hd_num, p_hd_num.text, size=FONT_SIZE_DEFAULT)


    # 2. Tên Hợp đồng
    p_tenloai = document.add_paragraph(f"HỢP ĐỒNG {contract_type.upper()}")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_tenloai, p_tenloai.text, size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Căn cứ ký hợp đồng (nếu có)
    # Cần tách phần căn cứ khỏi body hoặc nhận từ data
    can_cu_section = data.get("can_cu", None)
    if can_cu_section:
         if isinstance(can_cu_section, list):
             for can_cu_line in can_cu_section:
                  p_cc = document.add_paragraph(can_cu_line)
                  set_paragraph_format(p_cc, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(0), line_spacing=1.0)
                  add_run_with_format(p_cc, p_cc.text, size=FONT_SIZE_DEFAULT, italic=True)
         else:
              p_cc = document.add_paragraph(can_cu_section)
              set_paragraph_format(p_cc, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
              add_run_with_format(p_cc, p_cc.text, size=FONT_SIZE_DEFAULT, italic=True)
         document.add_paragraph() # Thêm khoảng trống


    # 4. Thông tin các bên
    p_date_place = document.add_paragraph(f"Hôm nay, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}, tại {data.get('signing_location', '...................................')}, chúng tôi gồm có:")
    set_paragraph_format(p_date_place, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    # Bên A
    p_party_a_name = document.add_paragraph()
    set_paragraph_format(p_party_a_name, space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_party_a_name, party_a_info['name'], bold=True)
    for detail in party_a_info['details']:
        p_detail = document.add_paragraph()
        set_paragraph_format(p_detail, left_indent=Cm(0.5), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_detail, detail, size=FONT_SIZE_DEFAULT)
    # Bên B
    p_party_b_name = document.add_paragraph()
    set_paragraph_format(p_party_b_name, space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_party_b_name, party_b_info['name'], bold=True)
    for detail in party_b_info['details']:
        p_detail = document.add_paragraph()
        set_paragraph_format(p_detail, left_indent=Cm(0.5), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_detail, detail, size=FONT_SIZE_DEFAULT)

    document.add_paragraph("Sau khi bàn bạc, hai bên thống nhất ký kết Hợp đồng với các điều khoản sau:", space_before=Pt(12))


    # 5. Các điều khoản hợp đồng
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            is_dieu = stripped_line.upper().startswith("ĐIỀU")
            is_khoan = re.match(r'^\d+\.\s+', stripped_line) # 1. 2.
            is_diem = re.match(r'^[a-z]\)\s+', stripped_line) # a) b)

            align = WD_ALIGN_PARAGRAPH.LEFT if is_dieu else WD_ALIGN_PARAGRAPH.JUSTIFY
            left_indent = Cm(0)
            first_indent = FIRST_LINE_INDENT

            if is_dieu:
                first_indent = Cm(0)
                align = WD_ALIGN_PARAGRAPH.CENTER # Điều thường căn giữa hoặc trái+đậm
            elif is_khoan:
                left_indent = Cm(0.5)
                first_indent = Cm(0)
            elif is_diem:
                left_indent = Cm(1.0)
                first_indent = Cm(0)

            set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_dieu)

    # 6. Điều khoản thi hành, hiệu lực
    # (Thường là điều cuối cùng trong body)

    # 7. Chữ ký các bên (Dùng table 2 cột)
    document.add_paragraph() # Khoảng cách trước chữ ký
    sig_table = document.add_table(rows=1, cols=2)
    sig_table.autofit = False
    sig_table.columns[0].width = Inches(3.0)
    sig_table.columns[1].width = Inches(3.0)

    # Chữ ký Bên A
    cell_a = sig_table.cell(0, 0)
    cell_a._element.clear_content()
    p_a_title = cell_a.add_paragraph("ĐẠI DIỆN BÊN A") # Hoặc lấy tên Bên A
    set_paragraph_format(p_a_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_a_title, p_a_title.text, size=FONT_SIZE_SIGNATURE, bold=True)
    # Lấy tên người đại diện từ details nếu có
    signer_a_name = party_a_info.get('representative', ' ')
    if 'Đại diện bởi:' in party_a_info['details'][3]: # Giả định vị trí
         signer_a_name = party_a_info['details'][3].split(':')[-1].strip()

    p_a_space = cell_a.add_paragraph("\n\n\n\n") # Chừa chỗ ký
    set_paragraph_format(p_a_space, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    p_a_signer = cell_a.add_paragraph(signer_a_name)
    set_paragraph_format(p_a_signer, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_run_with_format(p_a_signer, signer_a_name, size=FONT_SIZE_SIGNER_NAME, bold=True)


    # Chữ ký Bên B
    cell_b = sig_table.cell(0, 1)
    cell_b._element.clear_content()
    p_b_title = cell_b.add_paragraph("ĐẠI DIỆN BÊN B")
    set_paragraph_format(p_b_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_b_title, p_b_title.text, size=FONT_SIZE_SIGNATURE, bold=True)
    # Lấy tên người đại diện từ details nếu có
    signer_b_name = party_b_info.get('representative', ' ')
    if 'Đại diện bởi:' in party_b_info['details'][3]: # Giả định vị trí
         signer_b_name = party_b_info['details'][3].split(':')[-1].strip()

    p_b_space = cell_b.add_paragraph("\n\n\n\n") # Chừa chỗ ký
    set_paragraph_format(p_b_space, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    p_b_signer = cell_b.add_paragraph(signer_b_name)
    set_paragraph_format(p_b_signer, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_run_with_format(p_b_signer, signer_b_name, size=FONT_SIZE_SIGNER_NAME, bold=True)


    print("Định dạng Hợp đồng hoàn tất.")
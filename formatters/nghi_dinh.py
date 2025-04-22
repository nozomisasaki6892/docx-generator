# formatters/nghi_dinh.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, \
                   FONT_SIZE_HEADER, FONT_SIZE_SIGNATURE, FONT_SIZE_SIGNER_NAME, \
                   FONT_SIZE_PLACE_DATE, FONT_SIZE_DOC_NUMBER
try:
    from .common_elements import format_recipient_list
except ImportError:
    from common_elements import format_recipient_list

def format_nghi_dinh_header(document, issuer_name):
    p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_qh, p_qh.text, size=Pt(12), bold=True)
    p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True)
    p_line_tn = document.add_paragraph("-" * 20)
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    p_issuer = document.add_paragraph(issuer_name.upper())
    set_paragraph_format(p_issuer, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_issuer, p_issuer.text, size=Pt(13), bold=True)

def format_nghi_dinh_signature(document, signer_title, signer_name):
     sig_paragraph = document.add_paragraph()
     set_paragraph_format(sig_paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(6), space_after=Pt(0), line_spacing=1.15)
     add_run_with_format(sig_paragraph, "TM. CHÍNH PHỦ\n", size=FONT_SIZE_SIGNATURE, bold=True)
     add_run_with_format(sig_paragraph, signer_title.upper() + "\n\n\n\n\n", size=FONT_SIZE_SIGNATURE, bold=True)
     add_run_with_format(sig_paragraph, signer_name, size=FONT_SIZE_SIGNER_NAME, bold=True)


def format(document, data):
    print("Bắt đầu định dạng Nghị định (hành chính)...")
    title = data.get("title", "Nghị định về ABC")
    body = data.get("body", "Căn cứ...\nChính phủ ban hành Nghị định...\nĐiều 1...\nĐiều 2...")
    decree_number = data.get("decree_number", "Số: .../NĐ-CP")
    issuing_date_str = data.get("issuing_date", time.strftime("ngày %d tháng %m năm %Y"))
    issuing_location = data.get("issuing_location", "Hà Nội")
    issuer_name = data.get("issuing_org", "CHÍNH PHỦ").upper()

    # 1. Header (CHÍNH PHỦ)
    format_nghi_dinh_header(document, issuer_name)

    # 2. Số hiệu và Ngày tháng ban hành (Căn giữa dưới Header)
    p_num_date = document.add_paragraph()
    set_paragraph_format(p_num_date, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    # Thêm Số trước
    run_num = add_run_with_format(p_num_date, f"{decree_number}       ", size=FONT_SIZE_DEFAULT) # Cỡ 14
    # Thêm Ngày tháng sau, căn phải bằng khoảng trắng hoặc Tab (khó hơn)
    run_date = add_run_with_format(p_num_date, f"{issuing_location}, {issuing_date_str}", size=FONT_SIZE_PLACE_DATE, italic=True) # Cỡ 14


    # 3. Tên loại NGHỊ ĐỊNH
    p_tenloai = document.add_paragraph("NGHỊ ĐỊNH")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
    add_run_with_format(p_tenloai, "NGHỊ ĐỊNH", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Trích yếu
    nd_title_extract = title.replace("Nghị định", "").strip()
    # Bỏ "về việc", "quy định",... nếu có ở đầu trích yếu
    if nd_title_extract.lower().startswith("về việc"):
        nd_title_extract = nd_title_extract.split(" ", 2)[-1]
    elif nd_title_extract.lower().startswith("quy định"):
         nd_title_extract = nd_title_extract.split(" ", 1)[-1]

    p_trichyeu = document.add_paragraph(nd_title_extract)
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_trichyeu, nd_title_extract, size=Pt(14), bold=True)
    p_line_ty = document.add_paragraph("-" * 15)
    set_paragraph_format(p_line_ty, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))


    # 5. Cơ quan ban hành (Lặp lại) - BỎ ĐI NẾU DÙNG HEADER Ở TRÊN
    # p_issuer_body = document.add_paragraph(issuer_name)
    # set_paragraph_format(p_issuer_body, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    # add_run_with_format(p_issuer_body, issuer_name, size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # 6. Căn cứ ban hành
    body_lines = body.split('\n')
    processed_indices = set()

    for i, line in enumerate(body_lines):
        stripped_line = line.strip()
        if not stripped_line: continue
        if stripped_line.lower().startswith("căn cứ") or "chính phủ ban hành nghị định" in stripped_line.lower():
            p = document.add_paragraph(stripped_line)
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(0))
            set_run_format(p.runs[0], size=FONT_SIZE_DEFAULT, italic=True)
            processed_indices.add(i)
            # Dừng sau câu "Chính phủ ban hành..."
            if "chính phủ ban hành nghị định" in stripped_line.lower():
                break
        # Dừng khi hết căn cứ
        elif any(l.strip().lower().startswith("căn cứ") for l in body_lines[:i]):
             break
    if processed_indices: document.add_paragraph() # Khoảng cách sau căn cứ


    # 7. Nội dung (Chương, Điều, Khoản, Điểm)
    for i, line in enumerate(body_lines):
        if i in processed_indices: continue
        stripped_line = line.strip()
        if not stripped_line: continue
        p = document.add_paragraph()

        is_chuong = stripped_line.upper().startswith("CHƯƠNG")
        is_dieu = stripped_line.upper().startswith("ĐIỀU")
        is_khoan = re.match(r'^\d+\.\s+', stripped_line)
        is_diem = re.match(r'^[a-z]\)\s+', stripped_line)

        align = WD_ALIGN_PARAGRAPH.JUSTIFY
        left_indent = Cm(0)
        first_indent = FIRST_LINE_INDENT if not (is_chuong or is_dieu or is_khoan or is_diem) else Cm(0)
        is_bold = False
        size = FONT_SIZE_DEFAULT # Cỡ chữ 13-14
        space_before = Pt(0)
        space_after = Pt(6)
        line_spacing = 1.5

        if is_chuong:
            align = WD_ALIGN_PARAGRAPH.CENTER
            is_bold = True
            size = Pt(14) # Cỡ chữ chương = cỡ chữ nội dung
            space_before = Pt(12)
        elif is_dieu:
            align = WD_ALIGN_PARAGRAPH.LEFT
            is_bold = True
            size = Pt(14) # Cỡ chữ điều = cỡ chữ nội dung
            space_before = Pt(6)
        elif is_khoan:
            left_indent = Cm(0.5)
        elif is_diem:
            left_indent = Cm(1.0)


        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=line_spacing, space_before=space_before, space_after=space_after)
        add_run_with_format(p, stripped_line, size=size, bold=is_bold)


    # 8. Chữ ký (TM. CHÍNH PHỦ, THỦ TƯỚNG)
    signer_title = data.get("signer_title", "THỦ TƯỚNG")
    signer_name = data.get("signer_name", "[Tên Thủ tướng]")
    format_nghi_dinh_signature(document, signer_title, signer_name)

    # 9. Nơi nhận (Nghị định hành chính có thể có nơi nhận)
    if data.get('recipients'):
        format_recipient_list(document, data)
    else:
        # Có thể thêm nơi nhận mặc định nếu muốn
        pass

    print("Định dạng Nghị định (hành chính) hoàn tất.")
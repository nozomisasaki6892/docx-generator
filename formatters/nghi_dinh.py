# formatters/nghi_dinh.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Import các hàm tiện ích và cấu hình
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, \
                   FONT_SIZE_HEADER, FONT_SIZE_SIGNATURE, FONT_SIZE_SIGNER_NAME, \
                   FONT_SIZE_PLACE_DATE, FONT_SIZE_DOC_NUMBER

# Import hàm tạo nơi nhận từ common_elements
try:
    from .common_elements import format_recipient_list
except ImportError:
    from common_elements import format_recipient_list


# Hàm tạo header riêng cho Nghị định (QH/TN, Tên CQ)
def format_nghị_định_header(document, issuer_name):
    p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_qh, p_qh.text, size=Pt(12), bold=True) # Cỡ QH 12-13
    p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True) # Cỡ TN 13-14
    p_line_tn = document.add_paragraph("-" * 20)
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    p_issuer = document.add_paragraph(issuer_name.upper())
    set_paragraph_format(p_issuer, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_issuer, p_issuer.text, size=Pt(13), bold=True) # Tên CQ ban hành 12-13, đậm

# Hàm tạo chữ ký riêng cho Nghị định (TM. CQ, Chức vụ)
def format_nghị_định_signature(document, signer_title, signer_name):
     sig_paragraph = document.add_paragraph()
     set_paragraph_format(sig_paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(6), space_after=Pt(0), line_spacing=1.15)
     # TM. CHÍNH PHỦ (Ví dụ)
     add_run_with_format(sig_paragraph, "TM. CHÍNH PHỦ\n", size=FONT_SIZE_SIGNATURE, bold=True)
     # THỦ TƯỚNG (Ví dụ)
     add_run_with_format(sig_paragraph, signer_title.upper() + "\n\n\n\n\n", size=FONT_SIZE_SIGNATURE, bold=True)
     # Tên Thủ tướng
     add_run_with_format(sig_paragraph, signer_name, size=FONT_SIZE_SIGNER_NAME, bold=True)


def format(document, data):
    print("Bắt đầu định dạng Nghị định (hành chính)...")
    title = data.get("title", "Nghị định về ABC")
    body = data.get("body", "Nội dung nghị định...")
    decree_number = data.get("decree_number", "Số: .../NĐ-CP") # Số của NĐ hành chính
    issuing_date_str = data.get("issuing_date", time.strftime("ngày %d tháng %m năm %Y"))
    issuing_location = data.get("issuing_location", "Hà Nội")

    # 1. Header (CHÍNH PHỦ - Ví dụ mặc định)
    format_nghị_định_header(document, "CHÍNH PHỦ")

    # 2. Số hiệu và Ngày tháng ban hành (Căn giữa dưới Header)
    p_num_date = document.add_paragraph(f"{decree_number}       {issuing_location}, {issuing_date_str}")
    set_paragraph_format(p_num_date, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    # Cần tách riêng để định dạng italic cho ngày tháng
    run_num = add_run_with_format(p_num_date, f"{decree_number}       ", size=FONT_SIZE_DEFAULT)
    run_date = add_run_with_format(p_num_date, f"{issuing_location}, {issuing_date_str}", size=FONT_SIZE_PLACE_DATE, italic=True)


    # 3. Tên loại NGHỊ ĐỊNH
    p_tenloai = document.add_paragraph("NGHỊ ĐỊNH")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
    add_run_with_format(p_tenloai, "NGHỊ ĐỊNH", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Trích yếu
    nd_title = title.replace("Nghị định", "").strip()
    p_trichyeu = document.add_paragraph(f"Về việc {nd_title}")
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_trichyeu, f"Về việc {nd_title}", size=Pt(14), bold=True) # Trích yếu đậm, cỡ 14
    p_line_ty = document.add_paragraph("-" * 15)
    set_paragraph_format(p_line_ty, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 5. Cơ quan ban hành (Lặp lại)
    p_issuer = document.add_paragraph("CHÍNH PHỦ")
    set_paragraph_format(p_issuer, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_issuer, p_issuer.text, size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # 6. Căn cứ ban hành (Italic, thụt lề)
    preamble = data.get("preamble", []) # Nên là list các dòng căn cứ
    if isinstance(preamble, str): preamble = preamble.split('\n') # Tách nếu là string

    if preamble:
        for line in preamble:
            stripped_line = line.strip()
            if stripped_line:
                p_pre = document.add_paragraph(stripped_line)
                set_paragraph_format(p_pre, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(0), line_spacing=1.5)
                add_run_with_format(p_pre, stripped_line, size=FONT_SIZE_DEFAULT, italic=True)
        document.add_paragraph() # Khoảng cách sau căn cứ

    # 7. Nội dung (Chương, Điều, Khoản, Điểm)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if not stripped_line: continue
        p = document.add_paragraph()

        is_chuong = stripped_line.upper().startswith("CHƯƠNG")
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
            size = Pt(13) # Cỡ chữ chương = cỡ chữ nội dung
            space_before = Pt(12)
        elif is_dieu:
            align = WD_ALIGN_PARAGRAPH.LEFT
            first_indent = Cm(0)
            is_bold = True
            size = Pt(13) # Cỡ chữ điều = cỡ chữ nội dung
            space_before = Pt(6)
        elif is_khoan:
            left_indent = Cm(0.5)
            first_indent = Cm(0)
        elif is_diem:
            left_indent = Cm(1.0)
            first_indent = Cm(0)

        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_before=space_before, space_after=space_after)
        add_run_with_format(p, stripped_line, size=size, bold=is_bold)

    # 8. Chữ ký (TM. CHÍNH PHỦ, THỦ TƯỚNG)
    format_nghị_định_signature(document, "THỦ TƯỚNG", data.get("signer_name", "[Tên Thủ tướng]"))

    # 9. Nơi nhận (Nghị định hành chính có thể có nơi nhận)
    if data.get('recipients'):
        format_recipient_list(document, data)
    else:
        # Có thể thêm nơi nhận mặc định nếu muốn
        pass

    print("Định dạng Nghị định (hành chính) hoàn tất.")
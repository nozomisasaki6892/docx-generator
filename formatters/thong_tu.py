# formatters/thong_tu.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
# Header Thông tư giống NĐ30 admin, Chữ ký Bộ trưởng/Thủ trưởng
from .common_elements import format_basic_header, format_signature_block, format_recipient_list
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Thông tư...")
    title = data.get("title", "Thông tư hướng dẫn ABC")
    body = data.get("body", "Nội dung thông tư...")
    circular_number = data.get("circular_number", "Số: .../20.../TT-B...") # VD: TT-BTC, TT-BGDĐT
    issuing_date_str = data.get("issuing_date", time.strftime("ngày %d tháng %m năm %Y"))
    issuing_location = data.get("issuing_location", "Hà Nội")
    # Cần tên Bộ/Cơ quan ban hành từ data
    issuing_org = data.get("issuing_org", "BỘ [TÊN BỘ]").upper()

    # 1. Header (Giống NĐ30)
    # Cần truyền đúng tên Bộ vào issuing_org
    data['issuing_org'] = issuing_org # Đảm bảo data có tên Bộ
    format_basic_header(document, data, "ThongTu")

    # 2. Tên loại THÔNG TƯ
    p_tenloai = document.add_paragraph("THÔNG TƯ")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "THÔNG TƯ", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Trích yếu
    tt_title = title.replace("Thông tư", "").strip()
    p_trichyeu = document.add_paragraph(tt_title)
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_trichyeu, tt_title, size=Pt(14), bold=True)
    p_line_ty = document.add_paragraph("-" * 15)
    set_paragraph_format(p_line_ty, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))


    # 4. Căn cứ ban hành
    preamble = data.get("preamble", "Căn cứ Nghị định số .../20.../NĐ-CP ngày ... tháng ... năm ... của Chính phủ quy định chức năng, nhiệm vụ, quyền hạn và cơ cấu tổ chức của Bộ ...;\nCăn cứ [Luật/Nghị định được hướng dẫn];\nTheo đề nghị của [Vụ trưởng/Cục trưởng ...];\nBộ trưởng Bộ ... ban hành Thông tư ...")
    preamble_lines = preamble.split('\n')
    for line in preamble_lines:
         p_pre = document.add_paragraph(line)
         set_paragraph_format(p_pre, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(0), line_spacing=1.5)
         add_run_with_format(p_pre, line, size=FONT_SIZE_DEFAULT, italic=True)
    document.add_paragraph()

    # 5. Nội dung (Chương, Mục, Điều, Khoản, Điểm) - Tương tự Luật/Nghị định
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

    # 6. Chữ ký (BỘ TRƯỞNG/THỨ TRƯỞNG)
    # Cần lấy đúng chức vụ ký từ data
    if not data.get('signer_title'): data['signer_title'] = "BỘ TRƯỞNG" # Hoặc KT. BỘ TRƯỞNG, THỨ TRƯỞNG
    format_signature_block(document, data) # Dùng format chữ ký chuẩn NĐ30

    # 7. Nơi nhận (Khá nhiều nơi theo quy định)
    default_recipients_tt = [
        "- Văn phòng Chính phủ;",
        "- Các Bộ, cơ quan ngang Bộ, cơ quan thuộc Chính phủ;",
        "- UBND các tỉnh, thành phố trực thuộc Trung ương;",
        "- Cục Kiểm tra văn bản QPPL (Bộ Tư pháp);",
        "- Công báo;",
        "- Website Chính phủ;",
        f"- Website {issuing_org};",
        "- Lưu: VT, [Đơn vị soạn thảo]."
    ]
    if not data.get('recipients'): data['recipients'] = default_recipients_tt
    format_recipient_list(document, data)

    print("Định dạng Thông tư hoàn tất.")
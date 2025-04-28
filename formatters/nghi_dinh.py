# formatters/nghi_dinh.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list, add_national_emblem
except ImportError:
    # Fallback if running this file independently
    def format_basic_header(document, data, doc_type): pass
    def format_signature_block(document, data): pass
    def format_recipient_list(document, data): pass
    def add_national_emblem(document): pass


from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_VV, DEFAULT_LINE_SPACING

def format(document, data):
    print("Bắt đầu định dạng Nghị định...")

    # Data extraction with defaults
    issuance_number = data.get("issuance_number", "...../20../NĐ-CP")
    issuance_place = data.get("issuance_place", "Hà Nội")
    issuance_date = data.get("issuance_date", f"ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    title = data.get("title", "Nghị định về việc...")
    can_cu_list = data.get("can_cu", ["Căn cứ [Tên Luật/Pháp lệnh/Nghị quyết...];", "Căn cứ [Tên Nghị định...];", "Theo đề nghị của [Tên Bộ/Cơ quan];", "Chính phủ ban hành Nghị định quy định về [Nội dung]."])
    body_content = data.get("body", "Chương I\nNHỮNG QUY ĐỊNH CHUNG\nĐiều 1. Phạm vi điều chỉnh\n1. Nghị định này quy định về...\nĐiều 2. Đối tượng áp dụng\n...\nChương II\n[TÊN CHƯƠNG II]\nĐiều...\n...")
    signer_title = data.get("signer_title", "KT. THỦ TƯỚNG\nPHÓ THỦ TƯỚNG")
    signer_name = data.get("signer_name", "[Họ và tên]")

    # 1. Header (Quốc hiệu, Tiêu ngữ, Cơ quan ban hành, Số, Ký hiệu, Địa danh, Ngày tháng)
    add_national_emblem(document) # Add emblem
    format_basic_header(document, data, "NghiDinh", include_agency=True, agency_name="CHÍNH PHỦ", include_issuance_info=True, number=issuance_number, doc_type_text="Nghị định", place=issuance_place, date=issuance_date)

    # 2. Title "NGHỊ ĐỊNH"
    p_nghidinh = document.add_paragraph("NGHỊ ĐỊNH")
    set_paragraph_format(p_nghidinh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    set_run_format(p_nghidinh.runs[0], size=FONT_SIZE_TITLE, bold=True)

    # 3. Subject Title
    p_subject = document.add_paragraph(title.upper())
    set_paragraph_format(p_subject, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(18))
    set_run_format(p_subject.runs[0], size=FONT_SIZE_TITLE, bold=True)


    # 4. Căn cứ
    for item in can_cu_list:
        p_cc = document.add_paragraph(item)
        set_paragraph_format(p_cc, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=DEFAULT_LINE_SPACING, space_after=Pt(0))
        set_run_format(p_cc.runs[0], size=FONT_SIZE_DEFAULT)

    document.add_paragraph() # Add a blank line after căn cứ

    # 5. Body Content (Chapters, Articles, Clauses, Points)
    lines = body_content.strip().split('\n')
    current_chapter = None
    current_article = None

    for i, line in enumerate(lines):
        stripped_line = line.strip()
        if not stripped_line: continue

        # Detect Chapter
        chapter_match = re.match(r'^(Chương\s+\w+)\s*(.*)', stripped_line, re.IGNORECASE)
        if chapter_match:
            chapter_number = chapter_match.group(1)
            chapter_title = chapter_match.group(2).strip().upper()
            p_chapter = document.add_paragraph(f"{chapter_number}\n{chapter_title}")
            set_paragraph_format(p_chapter, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
            add_run_with_format(p_chapter, p_chapter.runs[0].text, size=FONT_SIZE_DEFAULT, bold=True)
            current_chapter = chapter_number
            current_article = None # Reset article counter for new chapter
            continue

        # Detect Article
        article_match = re.match(r'^(Điều\s+\w+)\.\s*(.*)', stripped_line, re.IGNORECASE)
        if article_match:
            article_number = article_match.group(1)
            article_title = article_match.group(2).strip()
            p_article = document.add_paragraph()
            add_run_with_format(p_article, f"{article_number}. ", size=FONT_SIZE_DEFAULT, bold=True)
            add_run_with_format(p_article, article_title, size=FONT_SIZE_DEFAULT, bold=True)
            set_paragraph_format(p_article, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=DEFAULT_LINE_SPACING, space_before=Pt(6), space_after=Pt(0))
            current_article = article_number
            continue

        # Detect Clause (numbered list item)
        clause_match = re.match(r'^(\d+)\.\s*(.*)', stripped_line)
        if clause_match:
            clause_number = clause_match.group(1)
            clause_content = clause_match.group(2).strip()
            p_clause = document.add_paragraph()
            add_run_with_format(p_clause, f"{clause_number}. ", size=FONT_SIZE_DEFAULT)
            add_run_with_format(p_clause, clause_content, size=FONT_SIZE_DEFAULT)
            set_paragraph_format(p_clause, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0.7), first_line_indent=Cm(-0.7), line_spacing=DEFAULT_LINE_SPACING, space_after=Pt(0))
            continue

        # Detect Point (lettered list item)
        point_match = re.match(r'^([a-zA-Z])\)\s*(.*)', stripped_line)
        if point_match:
            point_letter = point_match.group(1)
            point_content = point_match.group(2).strip()
            p_point = document.add_paragraph()
            add_run_with_format(p_point, f"{point_letter}) ", size=FONT_SIZE_DEFAULT)
            add_run_with_format(p_point, point_content, size=FONT_SIZE_DEFAULT)
            set_paragraph_format(p_point, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(1.4), first_line_indent=Cm(-0.7), line_spacing=DEFAULT_LINE_SPACING, space_after=Pt(0)) # Adjust indents as needed
            continue

        # Regular paragraph
        p_body = document.add_paragraph(stripped_line)
        set_paragraph_format(p_body, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=DEFAULT_LINE_SPACING, space_after=Pt(0))
        set_run_format(p_body.runs[0], size=FONT_SIZE_DEFAULT)

    # 6. Signature Block
    document.add_paragraph() # Add space before signature
    format_signature_block(document, data, signer_title=signer_title, signer_name=signer_name)

    # 7. Nơi nhận
    if 'recipients' not in data: data['recipients'] = ["Nơi nhận:", "- Ban Bí thư Trung ương Đảng;", "- Thủ tướng, các Phó Thủ tướng Chính phủ;", "- Các bộ, cơ quan ngang bộ, cơ quan thuộc Chính phủ;", "- HĐND, UBND các tỉnh, thành phố trực thuộc Trung ương;", "- Văn phòng Trung ương và các Ban của Đảng;", "- Văn phòng Quốc hội, Văn phòng Chủ tịch nước;", "- Hội đồng Dân tộc và các Ủy ban của Quốc hội;", "- Kiểm toán nhà nước;", "- Tòa án nhân dân tối cao, Viện kiểm sát nhân dân tối cao;", "- Ủy ban Giám sát tài chính Quốc gia;", "- Ngân hàng Chính sách xã hội;", "- Ngân hàng Phát triển Việt Nam;", "- Ủy ban Trung ương Mặt trận Tổ quốc Việt Nam và cơ quan ở Trung ương của các tổ chức chính trị - xã hội;", "- Văn phòng Chính phủ:", "+ BTCN, các PCN;", "+ Cổng TTĐT Chính phủ;", "+ Vụ..., Cục..., Đơn vị trực thuộc...;", "+ Lưu: VT, [Hồ sơ]."]
    format_recipient_list(document, data)


    print("Định dạng Nghị định hoàn tất.")
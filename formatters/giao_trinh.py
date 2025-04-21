# formatters/giao_trinh.py
import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Giáo trình (cơ bản)...")
    title = data.get("title", "GIÁO TRÌNH [TÊN MÔN HỌC/LĨNH VỰC]")
    authors = data.get("authors", ["[Tên tác giả 1]", "[Tên tác giả 2]"])
    publisher = data.get("publisher", "Nhà xuất bản [Tên NXB]")
    year = data.get("year", time.strftime("%Y"))
    body = data.get("body", "Lời nói đầu...\nChương 1...\n...") # Nội dung chính của giáo trình

    # --- Trang bìa (Ví dụ đơn giản) ---
    p_title = document.add_paragraph(title.upper())
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(72), space_after=Pt(36))
    add_run_with_format(p_title, title.upper(), size=Pt(18), bold=True)

    p_authors = document.add_paragraph("Tác giả:\n" + "\n".join(authors))
    set_paragraph_format(p_authors, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(72))
    add_run_with_format(p_authors, p_authors.text, size=FONT_SIZE_DEFAULT, bold=True)

    p_publisher = document.add_paragraph(publisher)
    set_paragraph_format(p_publisher, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_publisher, publisher, size=FONT_SIZE_DEFAULT)

    p_year = document.add_paragraph(str(year))
    set_paragraph_format(p_year, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(72))
    add_run_with_format(p_year, str(year), size=FONT_SIZE_DEFAULT)

    document.add_page_break()

    # --- Nội dung Giáo trình ---
    # Xử lý Lời nói đầu, Mục lục (nếu có trong body)
    # Định dạng Chương, Mục, nội dung
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if not stripped_line:
            # Thêm khoảng trống nếu cần giữa các đoạn
            # document.add_paragraph() # Có thể thêm dòng trống
            continue

        p = document.add_paragraph()

        is_chapter = stripped_line.upper().startswith("CHƯƠNG")
        is_section_roman = re.match(r'^[IVXLCDM]+\.\s+', stripped_line) # I. II.
        is_section_arabic = re.match(r'^\d+\.\s+', stripped_line) # 1. 2.
        is_subsection_alpha = re.match(r'^[a-z]\)\s+', stripped_line) # a) b)
        is_subsubsection_dash = stripped_line.startswith('-') # -

        align = WD_ALIGN_PARAGRAPH.JUSTIFY
        left_indent = Cm(0)
        first_indent = FIRST_LINE_INDENT
        is_bold = False
        size = FONT_SIZE_DEFAULT
        space_before = Pt(0)
        space_after = Pt(6)
        keep_with_next = False # Giữ tiêu đề với đoạn sau

        if is_chapter:
            align = WD_ALIGN_PARAGRAPH.CENTER
            first_indent = Cm(0)
            is_bold = True
            size = Pt(14) # Chương to hơn
            space_before = Pt(18) # Cách xa đoạn trước
            space_after = Pt(12)
            keep_with_next = True
            document.add_page_break() # Bắt đầu chương mới sang trang mới (tùy chọn)
        elif is_section_roman:
            align = WD_ALIGN_PARAGRAPH.LEFT
            first_indent = Cm(0)
            is_bold = True
            size = Pt(13) # Mục La Mã
            space_before = Pt(12)
            keep_with_next = True
        elif is_section_arabic:
            align = WD_ALIGN_PARAGRAPH.LEFT
            left_indent = Cm(0.5) # Thụt lề mục 1.
            first_indent = Cm(0)
            is_bold = True # Mục 1. đậm
            space_before = Pt(6)
            keep_with_next = True
        elif is_subsection_alpha:
            align = WD_ALIGN_PARAGRAPH.LEFT
            left_indent = Cm(1.0) # Thụt lề mục a)
            first_indent = Cm(0)
            is_bold = False # Thường không đậm
            # Có thể in nghiêng nếu muốn: italic=True
            space_before = Pt(3)
        elif is_subsubsection_dash:
            align = WD_ALIGN_PARAGRAPH.JUSTIFY
            left_indent = Cm(1.5) # Thụt lề mục -
            first_indent = Cm(0)
            space_before = Pt(0)
            space_after=Pt(3)


        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_before=space_before, space_after=space_after, keep_with_next=keep_with_next)
        add_run_with_format(p, stripped_line, size=size, bold=is_bold)

    print("Định dạng Giáo trình (cơ bản) hoàn tất.")
    print("LƯU Ý: Định dạng này chỉ xử lý cấu trúc Chương/Mục và text cơ bản.")
# formatters/tieu_luan.py
import re
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_NAME, FONT_SIZE_HEADER

def format_cover_page(document, data):
    school_info = data.get("school_info", {"university": "TÊN TRƯỜNG ĐẠI HỌC", "faculty": "KHOA [TÊN KHOA]"})
    student_info = data.get("student_info", {"name": "[Họ tên SV]", "id": "[Mã SV]", "class": "[Lớp]"})
    topic = data.get("title", "TIỂU LUẬN MÔN HỌC ABC")
    instructor = data.get("instructor", "[GV Hướng dẫn]")
    location = data.get("issuing_location", "Hà Nội")
    year = data.get("year", time.strftime("%Y"))

    # Canh giữa toàn bộ trang bìa
    # Tên trường, khoa
    p_uni = document.add_paragraph(school_info['university'].upper())
    set_paragraph_format(p_uni, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_uni, p_uni.text, size=FONT_SIZE_HEADER, bold=True)
    p_faculty = document.add_paragraph(school_info['faculty'].upper())
    set_paragraph_format(p_faculty, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_faculty, p_faculty.text, size=FONT_SIZE_HEADER, bold=True)
    document.add_paragraph("---------***---------", alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(36))

    # Tên loại TIỂU LUẬN
    p_type = document.add_paragraph("TIỂU LUẬN")
    set_paragraph_format(p_type, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_type, "TIỂU LUẬN", size=Pt(20), bold=True) # Cỡ lớn

    # Tên đề tài
    p_topic = document.add_paragraph(topic.upper())
    set_paragraph_format(p_topic, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(48))
    add_run_with_format(p_topic, topic.upper(), size=Pt(16), bold=True)

    # Thông tin GV và SV (Có thể dùng table ẩn hoặc canh lề)
    # Ví dụ dùng paragraph căn phải
    p_instructor = document.add_paragraph(f"Giáo viên hướng dẫn: {instructor}")
    set_paragraph_format(p_instructor, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=Inches(3.0), space_after=Pt(6))

    p_student_name = document.add_paragraph(f"Sinh viên thực hiện: {student_info['name']}")
    set_paragraph_format(p_student_name, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=Inches(3.0), space_after=Pt(0))
    p_student_id = document.add_paragraph(f"Mã số sinh viên: {student_info['id']}")
    set_paragraph_format(p_student_id, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=Inches(3.0), space_before=Pt(0), space_after=Pt(0))
    p_student_class = document.add_paragraph(f"Lớp: {student_info['class']}")
    set_paragraph_format(p_student_class, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=Inches(3.0), space_before=Pt(0), space_after=Pt(72))


    # Địa danh, năm
    p_loc_year = document.add_paragraph(f"{location}, năm {year}")
    set_paragraph_format(p_loc_year, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(36))
    add_run_with_format(p_loc_year, p_loc_year.text, size=FONT_SIZE_DEFAULT)

    document.add_page_break()


def format(document, data):
    print("Bắt đầu định dạng Tiểu luận...")
    body = data.get("body", "Lời mở đầu...\nChương 1...\n...")

    # 1. Tạo trang bìa
    format_cover_page(document, data)

    # 2. Định dạng nội dung (Lời mở đầu, Mục lục, Chương, Mục, Kết luận, TLTK)
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if not stripped_line: continue
        p = document.add_paragraph()

        # Nhận diện tiêu đề mục lớn (Thường in hoa, đậm, căn giữa)
        is_main_section_title = any(s in stripped_line.upper() for s in ["LỜI MỞ ĐẦU", "MỤC LỤC", "DANH MỤC", "PHẦN MỞ ĐẦU", "PHẦN NỘI DUNG", "KẾT LUẬN", "TÀI LIỆU THAM KHẢO"])

        is_chapter = stripped_line.upper().startswith("CHƯƠNG") # Chương 1: TỔNG QUAN...
        is_section_roman = re.match(r'^[IVXLCDM]+\.\s+', stripped_line) # I. II. ...
        is_section_arabic = re.match(r'^\d+\.\s+', stripped_line)      # 1. 2. ...
        is_subsection_arabic = re.match(r'^\d+\.\d+\.\s+', stripped_line) # 1.1. 1.2. ...
        is_subsubsection_arabic = re.match(r'^\d+\.\d+\.\d+\.\s+', stripped_line) # 1.1.1. ...
        is_subsection_alpha = re.match(r'^[a-z]\)\s+', stripped_line) # a) b) ...

        align = WD_ALIGN_PARAGRAPH.JUSTIFY
        left_indent = Cm(0)
        first_indent = FIRST_LINE_INDENT
        is_bold = False
        is_italic = False
        size = Pt(13) # Cỡ chữ phổ biến cho tiểu luận
        space_before = Pt(0)
        space_after = Pt(6)
        keep_with_next = False

        if is_main_section_title:
            align = WD_ALIGN_PARAGRAPH.CENTER
            first_indent = Cm(0)
            is_bold = True
            size = Pt(14)
            space_before = Pt(18)
            space_after = Pt(12)
            keep_with_next = True
        elif is_chapter:
            align = WD_ALIGN_PARAGRAPH.CENTER # Hoặc LEFT
            first_indent = Cm(0)
            is_bold = True
            size = Pt(14)
            space_before = Pt(12)
            space_after = Pt(12)
            keep_with_next = True
        elif is_section_roman or is_section_arabic: # Mục cấp 1 (I. hoặc 1.)
            align = WD_ALIGN_PARAGRAPH.LEFT
            first_indent = Cm(0)
            is_bold = True
            size = Pt(13)
            space_before = Pt(12)
            keep_with_next = True
        elif is_subsection_arabic: # Mục cấp 2 (1.1.)
            align = WD_ALIGN_PARAGRAPH.LEFT
            left_indent = Cm(0.5)
            first_indent = Cm(0)
            is_bold = True # Mục cấp 2 có thể đậm
            space_before = Pt(6)
            keep_with_next = True
        elif is_subsubsection_arabic: # Mục cấp 3 (1.1.1.)
            align = WD_ALIGN_PARAGRAPH.LEFT
            left_indent = Cm(1.0)
            first_indent = Cm(0)
            is_bold = False # Cấp 3 thường không đậm
            is_italic = True # Có thể in nghiêng
            space_before = Pt(6)
            keep_with_next = True
        elif is_subsection_alpha: # Mục a) b)
            align = WD_ALIGN_PARAGRAPH.JUSTIFY
            left_indent = Cm(1.5) # Thụt lề sâu hơn
            first_indent = Cm(0)
            space_before = Pt(3)


        set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_before=space_before, space_after=space_after, keep_with_next=keep_with_next)
        add_run_with_format(p, stripped_line, size=size, bold=is_bold, italic=is_italic)

    # Lưu ý: Cần có cơ chế tự động tạo Mục lục, đánh số trang (phức tạp hơn)

    print("Định dạng Tiểu luận hoàn tất.")
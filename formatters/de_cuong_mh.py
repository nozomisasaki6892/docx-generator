# formatters/de_cuong_mh.py
import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Đề cương môn học...")
    # Thông tin môn học
    course = data.get("course", {
        "name": "[TÊN MÔN HỌC]", "code": "[Mã HP]", "credits": "...",
        "prerequisites": "[Môn tiên quyết]", "department": "[Bộ môn phụ trách]", "faculty": "[Khoa]",
        "lecturer": "[Giảng viên]", "email": "[Email GV]"
    })
    # Nội dung đề cương (list các section, mỗi section có title và content)
    syllabus_content = data.get("syllabus_content", [
        {"title": "1. Thông tin chung về môn học", "content": ["Tên môn học:", "Mã môn học:", "Số tín chỉ:", "..."]},
        {"title": "2. Mục tiêu môn học", "content": ["Kiến thức:", "Kỹ năng:", "Thái độ:"]},
        {"title": "3. Chuẩn đầu ra môn học", "content": ["G1:", "G2:", "..."]},
        {"title": "4. Nội dung chi tiết môn học", "content": ["Tuần 1: [Nội dung]\n- [Chi tiết 1]\n- [Chi tiết 2]", "Tuần 2: [Nội dung]", "..."]},
        {"title": "5. Học liệu", "content": ["Giáo trình chính:", "Tài liệu tham khảo:", "..."]},
        {"title": "6. Đánh giá môn học", "content": ["Thành phần", "Trọng số", "Hình thức", "Điểm chuyên cần:", "Kiểm tra giữa kỳ:", "Thi cuối kỳ:"]}
    ])
    issuing_org = data.get("issuing_org", "TÊN TRƯỜNG").upper()
    faculty = course['faculty'].upper()

    # 1. Header (Tên trường, Khoa)
    p_org = document.add_paragraph(issuing_org)
    set_paragraph_format(p_org, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_org, issuing_org, size=FONT_SIZE_HEADER, bold=True)
    p_faculty = document.add_paragraph(f"KHOA {faculty}")
    set_paragraph_format(p_faculty, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_faculty, p_faculty.text, size=Pt(11), bold=True)

    # 2. Tên Đề cương
    p_title = document.add_paragraph("ĐỀ CƯƠNG CHI TIẾT MÔN HỌC")
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_title, p_title.text, size=FONT_SIZE_TITLE, bold=True, uppercase=True)
    p_course_name = document.add_paragraph(course['name'].upper())
    set_paragraph_format(p_course_name, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_course_name, p_course_name.text, size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Thông tin cơ bản môn học (ngoài phần nội dung)
    p_code = document.add_paragraph(f"Mã môn học: {course['code']}")
    set_paragraph_format(p_code, space_after=Pt(2))
    p_credits = document.add_paragraph(f"Số tín chỉ: {course['credits']}")
    set_paragraph_format(p_credits, space_before=Pt(2), space_after=Pt(2))
    p_prereq = document.add_paragraph(f"Môn học tiên quyết: {course['prerequisites']}")
    set_paragraph_format(p_prereq, space_before=Pt(2), space_after=Pt(12))


    # 4. Nội dung chi tiết đề cương
    for section in syllabus_content:
        section_title = section.get("title", "Nội dung")
        section_content = section.get("content", [])

        # Tiêu đề mục lớn (1., 2., ...)
        p_sec_title = document.add_paragraph(section_title)
        set_paragraph_format(p_sec_title, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(6), space_after=Pt(6))
        add_run_with_format(p_sec_title, section_title, size=FONT_SIZE_DEFAULT, bold=True)

        # Nội dung trong mục
        if isinstance(section_content, list):
            for item in section_content:
                 lines = item.split('\n') # Xử lý xuống dòng trong item
                 first_line = True
                 for line in lines:
                     stripped_line = line.strip()
                     if stripped_line:
                         p_item = document.add_paragraph()
                         # Thụt lề cho nội dung, gạch đầu dòng thụt sâu hơn
                         is_sub_item = stripped_line.startswith('-') or stripped_line.startswith('+')
                         left_indent = Cm(1.0) if is_sub_item else Cm(0.5)
                         first_indent = Cm(0) # Không thụt dòng đầu cho list/nội dung chi tiết
                         if first_line and not is_sub_item:
                             first_indent = FIRST_LINE_INDENT # Thụt dòng đầu cho đoạn văn đầu tiên

                         set_paragraph_format(p_item, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=left_indent, first_line_indent=first_indent, space_before=Pt(0), space_after=Pt(2), line_spacing=1.15)
                         add_run_with_format(p_item, stripped_line, size=FONT_SIZE_DEFAULT)
                     first_line = False
        elif isinstance(section_content, str): # Nếu content là 1 chuỗi lớn
             p_item = document.add_paragraph(section_content)
             set_paragraph_format(p_item, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0.5), first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6), line_spacing=1.5)
             add_run_with_format(p_item, section_content, size=FONT_SIZE_DEFAULT)


    # 5. Ngày tháng và Chữ ký (Trưởng bộ môn, Trưởng khoa)
    p_date_place_footer = document.add_paragraph(f"{data.get('issuing_location', '........')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_date_place_footer, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(12), space_after=Pt(0))
    add_run_with_format(p_date_place_footer, p_date_place_footer.text, size=FONT_SIZE_DEFAULT, italic=True)

    # Dùng table cho 2 chữ ký
    sig_table = document.add_table(rows=1, cols=2)
    sig_table.autofit = False
    sig_table.columns[0].width = Inches(3.0)
    sig_table.columns[1].width = Inches(3.0)

    # Chữ ký Trưởng bộ môn
    cell_tbm = sig_table.cell(0, 0)
    cell_tbm._element.clear_content()
    p_tbm_title = cell_tbm.add_paragraph("TRƯỞNG BỘ MÔN")
    set_paragraph_format(p_tbm_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_tbm_title, "TRƯỞNG BỘ MÔN", bold=True)
    # Tên Trưởng BM (nếu có)

    # Chữ ký Trưởng khoa
    cell_tk = sig_table.cell(0, 1)
    cell_tk._element.clear_content()
    p_tk_title = cell_tk.add_paragraph("TRƯỞNG KHOA")
    set_paragraph_format(p_tk_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_tk_title, "TRƯỞNG KHOA", bold=True)
    # Tên Trưởng khoa (nếu có)


    print("Định dạng Đề cương môn học hoàn tất.")
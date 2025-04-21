# formatters/giay_xac_nhan_sv.py
import re
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Dùng header, signature của trường/khoa
    from .common_elements import format_basic_header, format_signature_block
except ImportError:
    from common_elements import format_basic_header, format_signature_block
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Giấy xác nhận sinh viên...")
    # Thông tin sinh viên
    student = data.get("student", {"name": "[HỌ TÊN SINH VIÊN]", "dob": "__/__/____", "id": "[Mã số SV]", "class": "[Lớp]", "major": "[Ngành học]", "faculty": "[Khoa]", "course_year": "[Khóa học]", "status": "đang theo học"})
    purpose = data.get("purpose", "bổ túc hồ sơ vay vốn ngân hàng chính sách xã hội / xin tạm hoãn nghĩa vụ quân sự / ...") # Mục đích xác nhận
    issuing_org = data.get("issuing_org", "TÊN TRƯỜNG").upper()
    issuing_dept = data.get("issuing_dept", f"KHOA {student['faculty'].upper()}") # Khoa xác nhận

    # 1. Header (Tên trường và Tên Khoa) - Dùng table 2 cột
    header_table = document.add_table(rows=1, cols=2)
    header_table.autofit = False
    header_table.columns[0].width = Inches(3.0)
    header_table.columns[1].width = Inches(3.0)

    # Cột trái: Tên Trường
    cell_org = header_table.cell(0, 0)
    cell_org._element.clear_content()
    p_org = cell_org.add_paragraph(issuing_org)
    set_paragraph_format(p_org, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_org, issuing_org, size=FONT_SIZE_HEADER, bold=True)
    p_dept = cell_org.add_paragraph(issuing_dept)
    set_paragraph_format(p_dept, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_dept, issuing_dept, size=Pt(11), bold=True)
    p_line_org = cell_org.add_paragraph("-------***-------")
    set_paragraph_format(p_line_org, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))


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
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 2. Tên Giấy xác nhận
    p_tenloai = document.add_paragraph("GIẤY XÁC NHẬN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_tenloai, "GIẤY XÁC NHẬN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)
    add_run_with_format(p_tenloai, "\n(V/v: Xác nhận sinh viên)", size=FONT_SIZE_DEFAULT, bold=True) # Thêm V/v

    # 3. Nội dung xác nhận
    p_confirm_intro = document.add_paragraph()
    set_paragraph_format(p_confirm_intro, space_after=Pt(6))
    add_run_with_format(p_confirm_intro, f"{issuing_dept} xác nhận:")

    p_name = document.add_paragraph()
    set_paragraph_format(p_name, left_indent=Cm(1.0), space_after=Pt(0))
    add_run_with_format(p_name, "Họ và tên sinh viên:")
    add_run_with_format(p_name, f" {student['name']}", bold=True)

    p_dob = document.add_paragraph()
    set_paragraph_format(p_dob, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_dob, f"Ngày sinh: {student['dob']}")

    p_id = document.add_paragraph()
    set_paragraph_format(p_id, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_id, f"Mã số sinh viên: {student['id']}")

    p_class = document.add_paragraph()
    set_paragraph_format(p_class, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_class, f"Lớp: {student['class']}")

    p_major = document.add_paragraph()
    set_paragraph_format(p_major, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_major, f"Ngành học: {student['major']}")

    p_faculty = document.add_paragraph()
    set_paragraph_format(p_faculty, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_faculty, f"Thuộc Khoa: {student['faculty']}")

    p_course = document.add_paragraph()
    set_paragraph_format(p_course, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(6))
    add_run_with_format(p_course, f"Khóa học: {student['course_year']}")

    p_status = document.add_paragraph()
    set_paragraph_format(p_status, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_status, f"Hiện đang là sinh viên năm thứ ... hệ đào tạo ... hình thức đào tạo ... của Trường {issuing_org}.") # Cần điền thêm thông tin này
    add_run_with_format(p_status, f" Tình trạng: {student['status']}.")

    # 4. Mục đích xác nhận
    p_purpose = document.add_paragraph()
    set_paragraph_format(p_purpose, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_purpose, f"Giấy xác nhận này được cấp theo đề nghị của sinh viên {student['name']} để {purpose}.")

    p_validity = document.add_paragraph()
    set_paragraph_format(p_validity, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(12))
    add_run_with_format(p_validity, "Giấy này có giá trị trong vòng ... tháng kể từ ngày ký.")

    # 5. Chữ ký (Trưởng khoa hoặc người được ủy quyền)
    p_date_place_footer = document.add_paragraph(f"{data.get('issuing_location', 'Hà Nội')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_date_place_footer, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=Pt(0))
    add_run_with_format(p_date_place_footer, p_date_place_footer.text, size=FONT_SIZE_DEFAULT, italic=True)

    if not data.get('signer_title'): data['signer_title'] = f"TL. HIỆU TRƯỞNG\nTRƯỞNG KHOA {student['faculty'].upper()}" # Ví dụ
    format_signature_block(document, data)


    print("Định dạng Giấy xác nhận sinh viên hoàn tất.")
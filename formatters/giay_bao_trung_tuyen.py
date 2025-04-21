# formatters/giay_bao_trung_tuyen.py
import re
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Có thể dùng Header của trường
    from .common_elements import format_basic_header, format_signature_block
except ImportError:
    from common_elements import format_basic_header, format_signature_block
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Giấy báo trúng tuyển/nhập học...")
    # Thông tin thí sinh và trúng tuyển (cần từ data)
    student = data.get("student", {"name": "[HỌ TÊN THÍ SINH]", "dob": "__/__/____", "id": "[Số báo danh/CCCD]", "score": "...", "major": "[Ngành trúng tuyển]", "program_type": "Đại học chính quy"})
    enrollment_info = data.get("enrollment", {"time": "[Thời gian nhập học]", "location": "[Địa điểm nhập học]", "required_docs": ["Hồ sơ cần nộp 1", "Hồ sơ cần nộp 2"], "fee": "[Học phí/Kinh phí]"})
    issuing_org = data.get("issuing_org", "TÊN TRƯỜNG").upper()
    issuing_org_parent = data.get("issuing_org_parent", None)

    # 1. Header trường (Có thể dùng format_basic_header)
    data['issuing_org'] = issuing_org
    if issuing_org_parent: data['issuing_org_parent'] = issuing_org_parent
    # Dùng header căn trái cho trường
    # format_basic_header(document, data, "GiayBaoTrungTuyen")
    # Hoặc header đơn giản hơn
    p_org = document.add_paragraph(issuing_org)
    set_paragraph_format(p_org, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_org, issuing_org, size=FONT_SIZE_HEADER, bold=True)
    if issuing_org_parent:
         p_parent = document.add_paragraph(issuing_org_parent.upper())
         set_paragraph_format(p_parent, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(0), space_after=Pt(0))
         add_run_with_format(p_parent, issuing_org_parent, size=Pt(11))
    p_line_org = document.add_paragraph("-------***-------")
    set_paragraph_format(p_line_org, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))


    # 2. Ngày tháng, địa điểm
    p_date_place = document.add_paragraph(f"{data.get('issuing_location', 'Hà Nội')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_date_place, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=Pt(12))
    add_run_with_format(p_date_place, p_date_place.text, size=FONT_SIZE_DEFAULT, italic=True)


    # 3. Tên Giấy báo
    title_text = data.get("title", "GIẤY BÁO TRÚNG TUYỂN VÀ NHẬP HỌC")
    p_tenloai = document.add_paragraph(title_text.upper())
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_tenloai, title_text.upper(), size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Kính gửi (tên thí sinh)
    p_kg = document.add_paragraph(f"Kính gửi: Em {student['name']}")
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(6))
    add_run_with_format(p_kg, p_kg.text, size=FONT_SIZE_DEFAULT, bold=True)

    # 5. Nội dung thông báo trúng tuyển
    p_content1 = document.add_paragraph()
    set_paragraph_format(p_content1, first_line_indent=FIRST_LINE_INDENT)
    add_run_with_format(p_content1, f"Hội đồng tuyển sinh Trường {issuing_org} chúc mừng Em đã trúng tuyển vào học chương trình {student['program_type']} năm {time.strftime('%Y')} của Trường.")

    p_details = document.add_paragraph()
    set_paragraph_format(p_details, left_indent=Cm(1.0), space_after=Pt(6), line_spacing=1.15)
    add_run_with_format(p_details, f"- Họ và tên: {student['name']}\n", bold=True)
    add_run_with_format(p_details, f"- Ngày sinh: {student['dob']}\n")
    add_run_with_format(p_details, f"- Số báo danh/CCCD: {student['id']}\n")
    add_run_with_format(p_details, f"- Điểm xét tuyển: {student['score']}\n")
    add_run_with_format(p_details, f"- Ngành trúng tuyển: {student['major']}", bold=True)

    # 6. Thông tin nhập học
    p_enroll_intro = document.add_paragraph()
    set_paragraph_format(p_enroll_intro, first_line_indent=FIRST_LINE_INDENT, space_before=Pt(6))
    add_run_with_format(p_enroll_intro, "Để nhập học, Em cần chuẩn bị và thực hiện các thủ tục sau:")

    p_enroll_time = document.add_paragraph()
    set_paragraph_format(p_enroll_time, left_indent=Cm(1.0), space_after=Pt(0))
    add_run_with_format(p_enroll_time, f"1. Thời gian nhập học: {enrollment_info['time']}", bold=True)

    p_enroll_loc = document.add_paragraph()
    set_paragraph_format(p_enroll_loc, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_enroll_loc, f"2. Địa điểm nhập học: {enrollment_info['location']}")

    p_enroll_docs_label = document.add_paragraph()
    set_paragraph_format(p_enroll_docs_label, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_enroll_docs_label, "3. Hồ sơ nhập học cần nộp:")
    for doc in enrollment_info['required_docs']:
        p_doc_item = document.add_paragraph()
        set_paragraph_format(p_doc_item, left_indent=Cm(1.5), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_doc_item, f"- {doc}")

    p_enroll_fee = document.add_paragraph()
    set_paragraph_format(p_enroll_fee, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(12))
    add_run_with_format(p_enroll_fee, f"4. Kinh phí nhập học (tạm thu): {enrollment_info['fee']}")

    # Lời kết
    p_closing = document.add_paragraph("Đề nghị Em có mặt đầy đủ, đúng thời gian và địa điểm quy định.")
    set_paragraph_format(p_closing, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))

    # 7. Chữ ký (Hiệu trưởng/Chủ tịch HĐTS)
    if not data.get('signer_title'): data['signer_title'] = "HIỆU TRƯỞNG" # Hoặc tương đương
    format_signature_block(document, data)

    print("Định dạng Giấy báo trúng tuyển hoàn tất.")
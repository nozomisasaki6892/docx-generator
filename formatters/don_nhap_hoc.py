# formatters/don_nhap_hoc.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Đơn xin nhập học...")
    # Thông tin thí sinh/sinh viên
    student = data.get("student", {
        "name": "[HỌ TÊN]", "dob": "__/__/____", "pob": "[Nơi sinh]",
        "gender": "Nam/Nữ", "ethnicity": "[Dân tộc]", "religion": "[Tôn giáo]",
        "id_card": "[Số CMND/CCCD]", "id_date": "__/__/____", "id_place": "[Nơi cấp]",
        "phone": "[Số điện thoại]", "email": "[Địa chỉ email]",
        "address": "[Địa chỉ thường trú]",
        "high_school": "[Tên trường THPT]", "graduation_year": "...",
        "admission_method": "[Phương thức xét tuyển]", "major_registered": "[Ngành đăng ký]"
    })
    parent_info = data.get("parents", {"father_name": "[Họ tên cha]", "father_job": "[Nghề nghiệp]", "mother_name": "[Họ tên mẹ]", "mother_job": "[Nghề nghiệp]", "contact_address": "[Địa chỉ liên hệ]"})
    recipient = data.get("recipient", "Hội đồng tuyển sinh Trường [Tên trường]")

    # 1. Header (QH/TN)
    p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_qh, p_qh.text, size=FONT_SIZE_HEADER, bold=True)
    p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True)
    p_line_tn = document.add_paragraph("-" * 20)
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 2. Ngày tháng, địa điểm
    p_date_place = document.add_paragraph(f"{data.get('issuing_location', '........')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_date_place, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=Pt(12))
    add_run_with_format(p_date_place, p_date_place.text, size=FONT_SIZE_DEFAULT, italic=True)

    # 3. Tên đơn
    title_text = data.get("title", "ĐƠN XIN NHẬP HỌC")
    p_tenloai = document.add_paragraph(title_text.upper())
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(12))
    add_run_with_format(p_tenloai, title_text.upper(), size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Kính gửi
    p_kg = document.add_paragraph(f"Kính gửi: {recipient}")
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_kg, p_kg.text, size=FONT_SIZE_DEFAULT, bold=True)

    # 5. Thông tin cá nhân
    p_info_label = document.add_paragraph("I. THÔNG TIN CÁ NHÂN", space_after=Pt(2))
    set_run_format(p_info_label.runs[0], bold=True)
    p_name = document.add_paragraph(f"Họ và tên: {student['name'].upper()}", space_after=Pt(0))
    add_run_with_format(p_name.runs[0], size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)
    document.add_paragraph(f"Ngày sinh: {student['dob']}          Giới tính: {student['gender']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Nơi sinh: {student['pob']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Dân tộc: {student['ethnicity']}         Tôn giáo: {student['religion']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Số CMND/CCCD: {student['id_card']}    Ngày cấp: {student['id_date']}    Nơi cấp: {student['id_place']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Địa chỉ thường trú: {student['address']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Điện thoại: {student['phone']}         Email: {student['email']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Tốt nghiệp THPT tại trường: {student['high_school']} năm {student['graduation_year']}", space_after=Pt(6), space_before=Pt(0))

    # 6. Thông tin đăng ký nhập học
    p_reg_label = document.add_paragraph("II. THÔNG TIN ĐĂNG KÝ NHẬP HỌC", space_after=Pt(2))
    set_run_format(p_reg_label.runs[0], bold=True)
    document.add_paragraph(f"Đã trúng tuyển theo phương thức: {student['admission_method']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Đăng ký nhập học vào ngành: {student['major_registered']}", space_after=Pt(6), space_before=Pt(0))

    # 7. Thông tin gia đình
    p_fam_label = document.add_paragraph("III. THÔNG TIN GIA ĐÌNH", space_after=Pt(2))
    set_run_format(p_fam_label.runs[0], bold=True)
    document.add_paragraph(f"Họ tên cha: {parent_info['father_name']} - Nghề nghiệp: {parent_info['father_job']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Họ tên mẹ: {parent_info['mother_name']} - Nghề nghiệp: {parent_info['mother_job']}", space_after=Pt(0), space_before=Pt(0))
    document.add_paragraph(f"Địa chỉ liên hệ của gia đình: {parent_info['contact_address']}", space_after=Pt(6), space_before=Pt(0))

    # 8. Cam kết
    p_commit = document.add_paragraph("Tôi xin cam đoan những lời khai trên là đúng sự thật và xin chấp hành nghiêm chỉnh mọi quy chế, quy định của Nhà trường.", space_before=Pt(12), space_after=Pt(12), first_line_indent=FIRST_LINE_INDENT)

    # 9. Chữ ký (Người làm đơn và Phụ huynh) - Dùng table
    sig_table = document.add_table(rows=1, cols=2)
    sig_table.autofit = False
    sig_table.columns[0].width = Cm(8.0) / 2
    sig_table.columns[1].width = Cm(8.0) / 2

    # Chữ ký phụ huynh
    cell_parent = sig_table.cell(0, 0)
    cell_parent._element.clear_content()
    p_parent_title = cell_parent.add_paragraph("Ý KIẾN PHỤ HUYNH")
    set_paragraph_format(p_parent_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_parent_title, "Ý KIẾN PHỤ HUYNH", bold=True)
    p_parent_note = cell_parent.add_paragraph("(Ký, ghi rõ họ tên)")
    set_paragraph_format(p_parent_note, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_parent_note, "(Ký, ghi rõ họ tên)", size=Pt(11), italic=True)
    # p_parent_name = cell_parent.add_paragraph(parent_info['father_name'] or parent_info['mother_name'])
    # set_paragraph_format(p_parent_name, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    # add_run_with_format(p_parent_name, p_parent_name.text, bold=True)


    # Chữ ký người làm đơn
    cell_applicant = sig_table.cell(0, 1)
    cell_applicant._element.clear_content()
    p_app_title = cell_applicant.add_paragraph("NGƯỜI LÀM ĐƠN")
    set_paragraph_format(p_app_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_app_title, "NGƯỜI LÀM ĐƠN", bold=True)
    p_app_note = cell_applicant.add_paragraph("(Ký, ghi rõ họ tên)")
    set_paragraph_format(p_app_note, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_app_note, "(Ký, ghi rõ họ tên)", size=Pt(11), italic=True)
    p_app_name = cell_applicant.add_paragraph(student['name'])
    set_paragraph_format(p_app_name, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_run_with_format(p_app_name, student['name'], bold=True)

    print("Định dạng Đơn xin nhập học hoàn tất.")
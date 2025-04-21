# formatters/giay_nghi_phep.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header # Có thể dùng hoặc không
except ImportError:
    from common_elements import format_basic_header
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Giấy nghỉ phép...")
    # Thông tin cần thiết
    applicant = data.get("applicant", {"name": "[Họ tên]", "department": "[Bộ phận]", "position": "[Chức vụ]"})
    leave_days = data.get("leave_days", "...")
    start_date = data.get("start_date", "__/__/____")
    end_date = data.get("end_date", "__/__/____")
    reason = data.get("reason", "[Lý do xin nghỉ]")
    substitute = data.get("substitute", "[Người bàn giao công việc]")
    contact_address = data.get("contact_address", "[Địa chỉ liên hệ khi nghỉ]")
    contact_phone = data.get("contact_phone", "[Số điện thoại]")
    recipient_list = data.get("leave_recipients", ["Ban Giám đốc Công ty", "Phòng Hành chính - Nhân sự", "[Trưởng bộ phận]"])

    # 1. Header (Có thể chỉ là tên công ty hoặc có QH/TN)
    add_qh_tn = data.get("add_qh_tn_gnp", True)
    if add_qh_tn:
        p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_qh, p_qh.text, size=FONT_SIZE_HEADER, bold=True)
        p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
        set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True)
        p_line_tn = document.add_paragraph("-" * 20)
        set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    else: # Chỉ thêm tên công ty nếu không có QH/TN
        p_org = document.add_paragraph(data.get("issuing_org", "TÊN CÔNG TY").upper())
        set_paragraph_format(p_org, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
        add_run_with_format(p_org, p_org.text, size=FONT_SIZE_HEADER, bold=True)


    # 2. Ngày tháng làm đơn
    p_date_place = document.add_paragraph(f"{data.get('issuing_location', 'Hà Nội')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_date_place, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=Pt(12))
    add_run_with_format(p_date_place, p_date_place.text, size=FONT_SIZE_DEFAULT, italic=True)


    # 3. Tên đơn
    p_tenloai = document.add_paragraph("ĐƠN XIN NGHỈ PHÉP")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(12))
    add_run_with_format(p_tenloai, "ĐƠN XIN NGHỈ PHÉP", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 4. Kính gửi
    p_kg = document.add_paragraph()
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(6))
    add_run_with_format(p_kg, "Kính gửi:", bold=True)
    for recipient in recipient_list:
         p_rec = document.add_paragraph()
         set_paragraph_format(p_rec, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0))
         add_run_with_format(p_rec, f"- {recipient}", bold=True) # Kính gửi đậm

    # 5. Thông tin người làm đơn
    p_applicant_info = document.add_paragraph()
    set_paragraph_format(p_applicant_info, space_before=Pt(12), space_after=Pt(0))
    add_run_with_format(p_applicant_info, f"Tên tôi là: {applicant['name']}", bold=True)

    p_applicant_dept = document.add_paragraph()
    set_paragraph_format(p_applicant_dept, space_before=Pt(0), space_after=Pt(0))
    add_run_with_format(p_applicant_dept, f"Bộ phận công tác: {applicant['department']}")

    p_applicant_pos = document.add_paragraph()
    set_paragraph_format(p_applicant_pos, space_before=Pt(0), space_after=Pt(6))
    add_run_with_format(p_applicant_pos, f"Chức vụ: {applicant['position']}")

    # 6. Nội dung xin phép
    p_request = document.add_paragraph()
    set_paragraph_format(p_request, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_request, f"Nay tôi làm đơn này kính xin Ban lãnh đạo cho tôi được nghỉ phép {leave_days} ngày.")

    p_dates = document.add_paragraph()
    set_paragraph_format(p_dates, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_dates, f"Thời gian nghỉ: Từ ngày {start_date} đến hết ngày {end_date}.")

    p_reason = document.add_paragraph()
    set_paragraph_format(p_reason, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_reason, f"Lý do: {reason}")

    p_substitute = document.add_paragraph()
    set_paragraph_format(p_substitute, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_substitute, f"Trong thời gian nghỉ phép, công việc của tôi tại bộ phận đã được bàn giao lại cho Ông/Bà: {substitute}.")

    p_contact = document.add_paragraph()
    set_paragraph_format(p_contact, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(0))
    add_run_with_format(p_contact, f"Địa chỉ liên hệ trong thời gian nghỉ phép: {contact_address}")

    p_phone = document.add_paragraph()
    set_paragraph_format(p_phone, first_line_indent=FIRST_LINE_INDENT, space_before=Pt(0), space_after=Pt(12))
    add_run_with_format(p_phone, f"Số điện thoại: {contact_phone}")

    p_closing = document.add_paragraph()
    set_paragraph_format(p_closing, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(6))
    add_run_with_format(p_closing, "Kính mong Ban lãnh đạo xem xét và chấp thuận.")

    p_thanks = document.add_paragraph()
    set_paragraph_format(p_thanks, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(12))
    add_run_with_format(p_thanks, "Xin chân thành cảm ơn!")


    # 7. Chữ ký (Người làm đơn và các cấp duyệt) - Dùng table
    sig_table = document.add_table(rows=1, cols=3) # Ví dụ 3 cột: Người duyệt, HCNS, Người làm đơn
    sig_table.autofit = False
    col_width = Inches(2.0)
    sig_table.columns[0].width = col_width
    sig_table.columns[1].width = col_width
    sig_table.columns[2].width = col_width

    # Cột Người duyệt (ví dụ: Trưởng bộ phận)
    cell_approver1 = sig_table.cell(0, 0)
    cell_approver1._element.clear_content()
    p_app1_title = cell_approver1.add_paragraph(data.get("approver1_title", "Ý KIẾN TRƯỞNG BỘ PHẬN"))
    set_paragraph_format(p_app1_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_app1_title, p_app1_title.text, bold=True)

    # Cột HCNS (ví dụ)
    cell_hr = sig_table.cell(0, 1)
    cell_hr._element.clear_content()
    p_hr_title = cell_hr.add_paragraph(data.get("approver2_title", "Ý KIẾN PHÒNG HCNS"))
    set_paragraph_format(p_hr_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_hr_title, p_hr_title.text, bold=True)

    # Cột Người làm đơn
    cell_applicant = sig_table.cell(0, 2)
    cell_applicant._element.clear_content()
    p_app_title = cell_applicant.add_paragraph("NGƯỜI LÀM ĐƠN")
    set_paragraph_format(p_app_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_app_title, "NGƯỜI LÀM ĐƠN", bold=True)
    p_app_note = cell_applicant.add_paragraph("(Ký, ghi rõ họ tên)")
    set_paragraph_format(p_app_note, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_app_note, "(Ký, ghi rõ họ tên)", size=Pt(11), italic=True)
    p_app_name = cell_applicant.add_paragraph(applicant['name'])
    set_paragraph_format(p_app_name, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_run_with_format(p_app_name, applicant['name'], size=FONT_SIZE_SIGNER_NAME, bold=True)

    # Có thể thêm phần duyệt cuối cùng của Ban Giám đốc nếu cần

    print("Định dạng Giấy nghỉ phép hoàn tất.")
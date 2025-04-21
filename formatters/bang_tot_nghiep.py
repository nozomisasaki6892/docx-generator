# formatters/bang_tot_nghiep.py
import re
import time
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FONT_NAME, FONT_SIZE_HEADER, FONT_SIZE_SIGNATURE, FONT_SIZE_SIGNER_NAME

def format(document, data):
    print("Bắt đầu định dạng Bằng tốt nghiệp/Chứng chỉ (Cơ bản)...")
    # Thông tin cơ bản cần thiết
    degree_type = data.get("degree_type", "BẰNG TỐT NGHIỆP ĐẠI HỌC").upper() # HOẶC CHỨNG CHỈ...
    recipient = data.get("recipient", {"name": "[HỌ TÊN NGƯỜI NHẬN]", "dob": "__/__/____"})
    major = data.get("major", "[Ngành đào tạo]")
    degree_class = data.get("degree_class", "[Hạng tốt nghiệp]") # Xuất sắc, Giỏi, Khá, Trung bình
    mode_of_study = data.get("mode_of_study", "Chính quy") # Chính quy, Vừa làm vừa học,...
    conferral_decision_num = data.get("conferral_decision_num", "Số .../QĐ-...")
    conferral_date_str = data.get("conferral_date", time.strftime("ngày %d tháng %m năm %Y"))
    diploma_number = data.get("diploma_number", "Số vào sổ cấp bằng: ...")
    issuing_org = data.get("issuing_org", "TÊN TRƯỜNG").upper()
    issuing_org_parent = data.get("issuing_org_parent", None) # VD: BỘ GIÁO DỤC VÀ ĐÀO TẠO
    issuing_location = data.get("issuing_location", "Hà Nội")
    signer_title = data.get("signer_title", "HIỆU TRƯỞNG").upper()
    signer_name = data.get("signer_name", "[Tên Hiệu trưởng]")

    # ---- Cấu trúc văn bản cơ bản ----
    # (Bỏ qua các yếu tố đồ họa, chỉ tập trung vào text)

    # 1. Đơn vị cấp trên (nếu có) và Tên trường
    if issuing_org_parent:
        p_parent = document.add_paragraph(issuing_org_parent.upper())
        set_paragraph_format(p_parent, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(2))
        add_run_with_format(p_parent, p_parent.text, size=Pt(12), bold=False) # Cấp trên thường không đậm
    p_org = document.add_paragraph(issuing_org)
    set_paragraph_format(p_org, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_org, issuing_org, size=FONT_SIZE_HEADER, bold=True)

    # 2. Quốc hiệu / Tiêu ngữ
    p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_qh, p_qh.text, size=Pt(12), bold=True) # QH nhỏ hơn tên bằng
    p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True)
    p_line_tn = document.add_paragraph("-" * 15)
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(24))

    # 3. Tên Bằng/Chứng chỉ
    p_degree_title = document.add_paragraph(degree_type)
    set_paragraph_format(p_degree_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(18))
    # Cỡ chữ tên bằng thường rất lớn và font đặc biệt (khó tái tạo)
    add_run_with_format(p_degree_title, degree_type, size=Pt(24), bold=True) # Cỡ chữ lớn ví dụ

    # 4. Thông tin người nhận
    p_recipient_intro = document.add_paragraph("Chứng nhận Ông/Bà:") # Hoặc Cấp cho Ông/Bà
    set_paragraph_format(p_recipient_intro, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_recipient_intro, p_recipient_intro.text, size=FONT_SIZE_DEFAULT)

    p_recipient_name = document.add_paragraph(recipient['name'].upper())
    set_paragraph_format(p_recipient_name, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_recipient_name, p_recipient_name.text, size=Pt(16), bold=True, uppercase=True) # Tên người nhận to, đậm

    p_recipient_dob = document.add_paragraph(f"Sinh ngày: {recipient['dob']}")
    set_paragraph_format(p_recipient_dob, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_recipient_dob, p_recipient_dob.text, size=FONT_SIZE_DEFAULT)

    # 5. Nội dung công nhận
    p_major = document.add_paragraph(f"Đã tốt nghiệp ngành: {major}")
    set_paragraph_format(p_major, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_major, p_major.text, size=FONT_SIZE_DEFAULT, bold=True) # Ngành đậm

    p_class_mode = document.add_paragraph(f"Hạng tốt nghiệp: {degree_class} - Hình thức đào tạo: {mode_of_study}")
    set_paragraph_format(p_class_mode, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_class_mode, p_class_mode.text, size=FONT_SIZE_DEFAULT)

    # 6. Thông tin quyết định và ngày cấp
    p_decision = document.add_paragraph(f"Theo Quyết định số {conferral_decision_num} ngày {conferral_date_str}")
    set_paragraph_format(p_decision, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_decision, p_decision.text, size=FONT_SIZE_DEFAULT)

    p_issue_date = document.add_paragraph(f"{issuing_location}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_issue_date, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(12), space_after=Pt(0))
    add_run_with_format(p_issue_date, p_issue_date.text, size=FONT_SIZE_DEFAULT, italic=True)


    # 7. Chữ ký Hiệu trưởng
    sig_paragraph = document.add_paragraph()
    set_paragraph_format(sig_paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(6), space_after=Pt(0), line_spacing=1.15)
    add_run_with_format(sig_paragraph, signer_title + "\n\n\n\n\n", size=FONT_SIZE_SIGNATURE, bold=True)
    add_run_with_format(sig_paragraph, signer_name, size=FONT_SIZE_SIGNER_NAME, bold=True)


    # 8. Số vào sổ (thường góc dưới trái)
    p_diploma_num = document.add_paragraph(diploma_number)
    # Cần đặt vị trí tuyệt đối hoặc dùng text box/footer (phức tạp)
    # Tạm đặt ở cuối cùng, căn trái
    set_paragraph_format(p_diploma_num, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(12))
    add_run_with_format(p_diploma_num, diploma_number, size=Pt(10))

    print("Định dạng Bằng tốt nghiệp (cơ bản) hoàn tất.")
    print("LƯU Ý: Định dạng này chỉ chứa text, không bao gồm các yếu tố đồ họa của bằng thật.")
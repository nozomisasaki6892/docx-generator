# formatters/phieu_trinh.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Phiếu trình (Không bảng)...")
    # Thông tin cần từ data
    issuing_org = data.get("issuing_org", "TÊN ĐƠN VỊ TRÌNH").upper()
    issuing_dept = data.get("issuing_dept", None) # Bộ phận trình
    recipient_name = data.get("recipient_name", "[Tên lãnh đạo nhận]")
    issue_summary = data.get("issue_summary", "[Tóm tắt vấn đề trình bày]")
    proposal = data.get("proposal", "[Nội dung đề xuất/xin ý kiến]")
    attached_docs = data.get("attached_docs", None) # List hoặc string

    # 1. Header đơn giản (Đơn vị, Ngày tháng)
    p_org = document.add_paragraph(issuing_org)
    set_paragraph_format(p_org, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(0))
    add_run_with_format(p_org, issuing_org, size=FONT_SIZE_HEADER, bold=True)
    if issuing_dept:
         p_dept = document.add_paragraph(issuing_dept)
         set_paragraph_format(p_dept, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(6))
         add_run_with_format(p_dept, issuing_dept, size=Pt(11))

    p_date_place = document.add_paragraph(f"{data.get('issuing_location', '........')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_date_place, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_after=Pt(12))
    add_run_with_format(p_date_place, p_date_place.text, size=FONT_SIZE_DEFAULT, italic=True)

    # 2. Tên loại Phiếu trình
    p_tenloai = document.add_paragraph("PHIẾU TRÌNH GIẢI QUYẾT CÔNG VIỆC")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(12))
    add_run_with_format(p_tenloai, "PHIẾU TRÌNH GIẢI QUYẾT CÔNG VIỆC", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Kính gửi
    p_kg = document.add_paragraph(f"Kính gửi: {recipient_name}")
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(6))
    add_run_with_format(p_kg, p_kg.text, size=FONT_SIZE_DEFAULT, bold=True)

    # 4. Các mục nội dung
    p_issue_label = document.add_paragraph("1. Vấn đề trình:", space_after=Pt(0))
    add_run_with_format(p_issue_label, "1. Vấn đề trình:", bold=True)
    p_issue_content = document.add_paragraph(issue_summary, space_before=Pt(0), space_after=Pt(6), left_indent=Cm(0.5), first_line_indent=FIRST_LINE_INDENT)

    p_proposal_label = document.add_paragraph("2. Đề xuất/Xin ý kiến:", space_after=Pt(0))
    add_run_with_format(p_proposal_label, "2. Đề xuất/Xin ý kiến:", bold=True)
    p_proposal_content = document.add_paragraph(proposal, space_before=Pt(0), space_after=Pt(6), left_indent=Cm(0.5), first_line_indent=FIRST_LINE_INDENT)

    if attached_docs:
         p_attach_label = document.add_paragraph("3. Tài liệu kèm theo:", space_after=Pt(0))
         add_run_with_format(p_attach_label, "3. Tài liệu kèm theo:", bold=True)
         if isinstance(attached_docs, list):
             for doc in attached_docs:
                 p_doc = document.add_paragraph(f"- {doc}", space_before=Pt(0), space_after=Pt(0), left_indent=Cm(1.0))
         else:
              p_doc = document.add_paragraph(f"- {attached_docs}", space_before=Pt(0), space_after=Pt(0), left_indent=Cm(1.0))


    # 5. Ý kiến Lãnh đạo (Để trống)
    p_leader_label = document.add_paragraph("4. Ý kiến của Lãnh đạo:", space_before=Pt(12), space_after=Pt(60)) # Chừa nhiều dòng
    add_run_with_format(p_leader_label, "4. Ý kiến của Lãnh đạo:", bold=True)

    # 6. Chữ ký người trình
    p_signer_title = document.add_paragraph()
    set_paragraph_format(p_signer_title, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(12), space_after=Pt(60))
    add_run_with_format(p_signer_title, data.get("signer_title", "NGƯỜI TRÌNH").upper(), bold=True)

    p_signer_name = document.add_paragraph()
    set_paragraph_format(p_signer_name, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(0))
    add_run_with_format(p_signer_name, data.get("signer_name", "[Ký, ghi rõ họ tên]"), bold=True)


    print("Định dạng Phiếu trình (Không bảng) hoàn tất.")
# formatters/common_elements.py
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import (FONT_SIZE_HEADER, FONT_SIZE_DEFAULT, FONT_SIZE_SMALL,
                    FONT_SIZE_PLACE_DATE, FONT_SIZE_DOC_NUMBER, FONT_SIZE_SIGNATURE,
                    FONT_SIZE_SIGNER_NAME, FONT_SIZE_RECIPIENT_LABEL)

def format_basic_header(document, data, doc_type):
    issuing_org_parent = data.get("issuing_org_parent", None)
    issuing_org = data.get("issuing_org", "TÊN CƠ QUAN/TỔ CHỨC").upper()
    doc_number = data.get("doc_number", "Số:       /...")
    issuing_location = data.get("issuing_location", "Hà Nội")
    current_date_str = time.strftime(f"ngày %d tháng %m năm %Y")

    header_table = document.add_table(rows=1, cols=2)
    header_table.autofit = False
    header_table.columns[0].width = Inches(2.9) # Cột trái
    header_table.columns[1].width = Inches(3.3) # Cột phải (tổng ~6.2 inches < chiều rộng hiệu dụng)

    cell_org = header_table.cell(0, 0)
    cell_org._element.clear_content()
    align_org = WD_ALIGN_PARAGRAPH.CENTER

    if issuing_org_parent:
        p_org_parent = cell_org.add_paragraph(issuing_org_parent.upper())
        set_paragraph_format(p_org_parent, alignment=align_org, space_after=Pt(0))
        set_run_format(p_org_parent.runs[0], size=FONT_SIZE_HEADER, bold=False)

    p_org = cell_org.add_paragraph(issuing_org)
    set_paragraph_format(p_org, alignment=align_org, space_after=Pt(0))
    set_run_format(p_org.runs[0], size=FONT_SIZE_HEADER, bold=True)

    p_line_org = cell_org.add_paragraph("_______")
    set_paragraph_format(p_line_org, alignment=align_org, space_after=Pt(6))

    cell_qh_tn = header_table.cell(0, 1)
    cell_qh_tn._element.clear_content()
    p_qh = cell_qh_tn.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    set_run_format(p_qh.runs[0], size=FONT_SIZE_HEADER, bold=True)

    p_tn = cell_qh_tn.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    set_run_format(p_tn.runs[0], size=Pt(13), bold=True)

    p_line_tn = cell_qh_tn.add_paragraph("-" * 20)
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))

    # Table thứ 2 cho Số KH và Ngày tháng
    info_table = document.add_table(rows=1, cols=2)
    info_table.autofit = False
    info_table.columns[0].width = Inches(2.9)
    info_table.columns[1].width = Inches(3.3)

    cell_skh = info_table.cell(0, 0)
    cell_skh._element.clear_content()
    p_skh = cell_skh.add_paragraph(doc_number)
    set_paragraph_format(p_skh, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(0))
    set_run_format(p_skh.runs[0], size=FONT_SIZE_DOC_NUMBER)

    cell_ddnt = info_table.cell(0, 1)
    cell_ddnt._element.clear_content()
    p_ddnt = cell_ddnt.add_paragraph(f"{issuing_location}, {current_date_str}")
    set_paragraph_format(p_ddnt, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(0), space_after=Pt(6))
    set_run_format(p_ddnt.runs[0], size=FONT_SIZE_PLACE_DATE, italic=True)

    document.add_paragraph()

def format_signature_block(document, data):
    signer_title = data.get("signer_title", "CHỨC VỤ").upper()
    signer_name = data.get("signer_name", "Người Ký")
    authority_signer = data.get("authority_signer", None)

    sig_paragraph = document.add_paragraph()
    set_paragraph_format(sig_paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(6), space_after=Pt(0), line_spacing=1.15) # Căn phải, giãn dòng chữ ký hợp lý

    if authority_signer:
        add_run_with_format(sig_paragraph, authority_signer.upper() + "\n", size=FONT_SIZE_SIGNATURE, bold=True)

    add_run_with_format(sig_paragraph, signer_title + "\n", size=FONT_SIZE_SIGNATURE, bold=True)

    if data.get("signer_note"): # Thêm ghi chú dưới chức vụ nếu có
         add_run_with_format(sig_paragraph, f"({data.get('signer_note')})\n", size=Pt(11), italic=True) # Ghi chú nhỏ, nghiêng

    sig_paragraph.add_run("\n\n\n\n") # Khoảng trống ký

    add_run_with_format(sig_paragraph, signer_name, size=FONT_SIZE_SIGNER_NAME, bold=True)

def format_recipient_list(document, data):
    recipients = data.get("recipients", ["- Như trên;", "- Lưu: VT, ..."])

    p_nhan_label = document.add_paragraph()
    set_paragraph_format(p_nhan_label, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(12), space_after=Pt(0))
    add_run_with_format(p_nhan_label, "Nơi nhận:", size=FONT_SIZE_RECIPIENT_LABEL, bold=True, italic=True)

    for recipient in recipients:
        p_rec = document.add_paragraph()
        set_paragraph_format(p_rec, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=Cm(0.5), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_rec, recipient, size=FONT_SIZE_SMALL)
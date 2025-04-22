# formatters/common_elements.py
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import (FONT_SIZE_HEADER, FONT_SIZE_DEFAULT, FONT_SIZE_SMALL,
                    FONT_SIZE_PLACE_DATE, FONT_SIZE_DOC_NUMBER, FONT_SIZE_SIGNATURE,
                    FONT_SIZE_SIGNER_NAME, FONT_SIZE_RECIPIENT_LABEL)

def format_basic_header(document, data, doc_type):
    issuing_org_parent = data.get("issuing_org_parent", None)
    issuing_org = data.get("issuing_org", "TÊN CƠ QUAN/TỔ CHỨC").upper()
    doc_number = data.get("doc_number", "Số:       /...") # Thêm khoảng trắng để đẩy ra
    issuing_location = data.get("issuing_location", "Hà Nội")
    current_date_str = time.strftime(f"ngày %d tháng %m năm %Y")

    header_table = document.add_table(rows=1, cols=2)
    header_table.autofit = False
    header_table.allow_autofit = False
    # Điều chỉnh độ rộng để đẩy khối QH/TN và Ngày tháng sang phải
    header_table.columns[0].width = Inches(2.9)
    header_table.columns[1].width = Inches(3.3)

    # --- Cột trái: Cơ quan ban hành và Số hiệu ---
    cell_org = header_table.cell(0, 0)
    cell_org._element.clear_content()
    align_org_cell = WD_ALIGN_PARAGRAPH.CENTER

    # Cơ quan chủ quản (nếu có)
    if issuing_org_parent:
        p_org_parent = cell_org.add_paragraph(issuing_org_parent.upper())
        set_paragraph_format(p_org_parent, alignment=align_org_cell, space_after=Pt(0))
        set_run_format(p_org_parent.runs[0], size=FONT_SIZE_HEADER, bold=False) # Cỡ 12-13, không đậm

    # Tên cơ quan ban hành
    p_org = cell_org.add_paragraph(issuing_org)
    set_paragraph_format(p_org, alignment=align_org_cell, space_after=Pt(0))
    set_run_format(p_org.runs[0], size=FONT_SIZE_HEADER, bold=True) # Cỡ 12-13, đậm

    # Dấu gạch chân dưới tên CQBH
    p_line_org = cell_org.add_paragraph("_______") # Hoặc dùng shape nếu muốn đẹp hơn
    set_paragraph_format(p_line_org, alignment=align_org_cell, space_after=Pt(6))
    set_run_format(p_line_org.runs[0], size=FONT_SIZE_HEADER, bold=True)

    # Số hiệu văn bản (căn trái dưới dòng kẻ)
    p_skh = cell_org.add_paragraph(doc_number)
    # Căn lề trái trong ô nhưng vẫn thuộc cột trái của layout tổng thể
    set_paragraph_format(p_skh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(0))
    set_run_format(p_skh.runs[0], size=FONT_SIZE_DOC_NUMBER) # Cỡ 13

    # --- Cột phải: Quốc hiệu, Tiêu ngữ và Ngày tháng ---
    cell_qh_tn = header_table.cell(0, 1)
    cell_qh_tn._element.clear_content()

    # Quốc hiệu
    p_qh = cell_qh_tn.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    set_run_format(p_qh.runs[0], size=FONT_SIZE_HEADER, bold=True) # Cỡ 12-13, đậm

    # Tiêu ngữ
    p_tn = cell_qh_tn.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    set_run_format(p_tn.runs[0], size=Pt(13), bold=True) # Cỡ 13-14, đậm

    # Dấu gạch chân dưới tiêu ngữ
    p_line_tn = cell_qh_tn.add_paragraph("-" * 20) # Điều chỉnh độ dài gạch
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    set_run_format(p_line_tn.runs[0], size=Pt(13), bold=True)

    # Địa danh, ngày tháng (căn phải dưới dòng kẻ)
    p_ddnt = cell_qh_tn.add_paragraph(f"{issuing_location}, {current_date_str}")
    # Căn phải trong ô, thuộc cột phải của layout tổng thể
    set_paragraph_format(p_ddnt, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
    set_run_format(p_ddnt.runs[0], size=FONT_SIZE_PLACE_DATE, italic=True) # Cỡ 13-14, nghiêng

    # Thêm khoảng trắng sau header
    document.add_paragraph()


def format_signature_block(document, data):
    signer_title = data.get("signer_title", "CHỨC VỤ NGƯỜI KÝ").upper()
    signer_name = data.get("signer_name", "Người Ký")
    authority_signer = data.get("authority_signer", None) # VD: KT. BỘ TRƯỞNG, TM. ỦY BAN NHÂN DÂN
    signer_note = data.get("signer_note", None) # VD: (Đã ký)

    sig_paragraph = document.add_paragraph()
    # Căn phải toàn bộ khối chữ ký
    set_paragraph_format(sig_paragraph, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(6), space_after=Pt(0), line_spacing=1.0) # Giãn dòng đơn

    # Thẩm quyền ký (nếu có)
    if authority_signer:
        # Cần thêm khoảng trắng để đẩy chức vụ xuống dòng sau nếu thẩm quyền quá dài
        run_auth = add_run_with_format(sig_paragraph, authority_signer.upper() + "\n", size=FONT_SIZE_SIGNATURE, bold=True)

    # Chức vụ người ký
    run_title = add_run_with_format(sig_paragraph, signer_title + "\n", size=FONT_SIZE_SIGNATURE, bold=True)

    # Ghi chú dưới chức vụ (nếu có)
    if signer_note:
        run_note = add_run_with_format(sig_paragraph, f"{signer_note}\n", size=Pt(11), italic=True)

    # Khoảng trống ký tên (Thêm nhiều \n hơn)
    sig_paragraph.add_run("\n\n\n\n\n")

    # Tên người ký
    run_name = add_run_with_format(sig_paragraph, signer_name, size=FONT_SIZE_SIGNER_NAME, bold=True)

def format_recipient_list(document, data):
    recipients = data.get("recipients", [])
    # Cung cấp giá trị mặc định nếu recipients rỗng hoặc không có
    if not recipients:
        recipients = ["- Như trên;", "- Lưu: VT, ...;"]

    p_nhan_label = document.add_paragraph()
    set_paragraph_format(p_nhan_label, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(12), space_after=Pt(0))
    add_run_with_format(p_nhan_label, "Nơi nhận:", size=FONT_SIZE_RECIPIENT_LABEL, bold=True, italic=True) # Cỡ 12

    for recipient in recipients:
        p_rec = document.add_paragraph()
        # Thụt lề dòng đầu cho các mục nơi nhận
        set_paragraph_format(p_rec, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=Cm(0.7), first_line_indent=Cm(-0.7), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_rec, recipient, size=FONT_SIZE_SMALL) # Cỡ 11
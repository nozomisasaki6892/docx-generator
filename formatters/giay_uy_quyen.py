# formatters/giay_uy_quyen.py
import re
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header # Có thể dùng hoặc không
except ImportError:
    from common_elements import format_basic_header
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER, FONT_SIZE_SIGNATURE, FONT_SIZE_SIGNER_NAME

def format(document, data):
    print("Bắt đầu định dạng Giấy ủy quyền...")
    title = data.get("title", "Giấy ủy quyền")
    body = data.get("body", "Nội dung ủy quyền...")
    # Thông tin các bên (cần có cấu trúc rõ ràng từ data)
    principal = data.get("principal", {"name": "[Họ tên người ủy quyền]", "info": ["CMND/CCCD số:", "Địa chỉ:"]})
    agent = data.get("agent", {"name": "[Họ tên người được ủy quyền]", "info": ["CMND/CCCD số:", "Địa chỉ:"]})
    scope = data.get("scope", "[Nội dung công việc được ủy quyền]")
    duration = data.get("duration", "[Thời hạn ủy quyền, ví dụ: từ ngày ... đến ngày ... hoặc cho đến khi công việc hoàn thành]")

    # 1. Header (Có thể tùy chọn có QH/TN)
    add_qh_tn = data.get("add_qh_tn_guq", True)
    if add_qh_tn:
        p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
        set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_qh, p_qh.text, size=FONT_SIZE_HEADER, bold=True)
        p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
        set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        add_run_with_format(p_tn, p_tn.text, size=Pt(13), bold=True)
        p_line_tn = document.add_paragraph("-" * 20)
        set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 2. Tên loại
    p_tenloai = document.add_paragraph("GIẤY ỦY QUYỀN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(12))
    add_run_with_format(p_tenloai, "GIẤY ỦY QUYỀN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Ngày tháng địa điểm
    p_date_place = document.add_paragraph(f"{data.get('issuing_location', 'Hà Nội')}, ngày {time.strftime('%d')} tháng {time.strftime('%m')} năm {time.strftime('%Y')}")
    set_paragraph_format(p_date_place, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_date_place, p_date_place.text, size=FONT_SIZE_DEFAULT, italic=True)

    # 4. Thông tin người ủy quyền
    p_principal_label = document.add_paragraph()
    set_paragraph_format(p_principal_label, space_after=Pt(0))
    add_run_with_format(p_principal_label, "Bên ủy quyền (Bên A):", bold=True)
    p_principal_name = document.add_paragraph()
    set_paragraph_format(p_principal_name, left_indent=Cm(1.0), space_after=Pt(0))
    add_run_with_format(p_principal_name, f"Ông/Bà: {principal['name']}", bold=True)
    for info_line in principal['info']:
        p_info = document.add_paragraph()
        set_paragraph_format(p_info, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_info, info_line, size=FONT_SIZE_DEFAULT)

    # 5. Thông tin người được ủy quyền
    p_agent_label = document.add_paragraph()
    set_paragraph_format(p_agent_label, space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_agent_label, "Bên được ủy quyền (Bên B):", bold=True)
    p_agent_name = document.add_paragraph()
    set_paragraph_format(p_agent_name, left_indent=Cm(1.0), space_after=Pt(0))
    add_run_with_format(p_agent_name, f"Ông/Bà: {agent['name']}", bold=True)
    for info_line in agent['info']:
        p_info = document.add_paragraph()
        set_paragraph_format(p_info, left_indent=Cm(1.0), space_before=Pt(0), space_after=Pt(0), line_spacing=1.0)
        add_run_with_format(p_info, info_line, size=FONT_SIZE_DEFAULT)

    # 6. Nội dung ủy quyền
    p_scope_label = document.add_paragraph()
    set_paragraph_format(p_scope_label, space_before=Pt(12), space_after=Pt(0))
    add_run_with_format(p_scope_label, "Nội dung ủy quyền:", bold=True)
    p_scope_content = document.add_paragraph()
    set_paragraph_format(p_scope_content, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT)
    add_run_with_format(p_scope_content, scope, size=FONT_SIZE_DEFAULT)

    # 7. Thời hạn ủy quyền
    p_duration_label = document.add_paragraph()
    set_paragraph_format(p_duration_label, space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(p_duration_label, "Thời hạn ủy quyền:", bold=True)
    p_duration_content = document.add_paragraph()
    set_paragraph_format(p_duration_content, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT)
    add_run_with_format(p_duration_content, duration, size=FONT_SIZE_DEFAULT)

    # 8. Cam kết (nếu có)
    p_commitment = document.add_paragraph("Hai bên cam kết thực hiện đúng nội dung đã ủy quyền.", space_before=Pt(12))
    set_paragraph_format(p_commitment)

    # 9. Chữ ký hai bên (Dùng table)
    sig_table = document.add_table(rows=1, cols=2)
    sig_table.autofit = False
    sig_table.columns[0].width = Inches(3.0)
    sig_table.columns[1].width = Inches(3.0)

    # Chữ ký Bên được ủy quyền (Bên B - thường bên trái)
    cell_b = sig_table.cell(0, 0)
    cell_b._element.clear_content()
    p_b_title = cell_b.add_paragraph("BÊN ĐƯỢC ỦY QUYỀN")
    set_paragraph_format(p_b_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_b_title, "BÊN ĐƯỢC ỦY QUYỀN", size=FONT_SIZE_SIGNATURE, bold=True)
    p_b_note = cell_b.add_paragraph("(Ký, ghi rõ họ tên)")
    set_paragraph_format(p_b_note, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_b_note, "(Ký, ghi rõ họ tên)", size=Pt(11), italic=True)
    p_b_name = cell_b.add_paragraph(agent['name'])
    set_paragraph_format(p_b_name, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_run_with_format(p_b_name, agent['name'], size=FONT_SIZE_SIGNER_NAME, bold=True)

    # Chữ ký Bên ủy quyền (Bên A - thường bên phải)
    cell_a = sig_table.cell(0, 1)
    cell_a._element.clear_content()
    p_a_title = cell_a.add_paragraph("BÊN ỦY QUYỀN")
    set_paragraph_format(p_a_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_a_title, "BÊN ỦY QUYỀN", size=FONT_SIZE_SIGNATURE, bold=True)
    p_a_note = cell_a.add_paragraph("(Ký, ghi rõ họ tên)")
    set_paragraph_format(p_a_note, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60))
    add_run_with_format(p_a_note, "(Ký, ghi rõ họ tên)", size=Pt(11), italic=True)
    p_a_name = cell_a.add_paragraph(principal['name'])
    set_paragraph_format(p_a_name, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_run_with_format(p_a_name, principal['name'], size=FONT_SIZE_SIGNER_NAME, bold=True)

    # Có thể thêm phần xác nhận của cơ quan công chứng nếu cần

    print("Định dạng Giấy ủy quyền hoàn tất.")
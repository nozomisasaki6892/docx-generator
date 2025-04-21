# formatters/quyet_dinh_ts.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    # Quyết định của trường dùng header, signature chuẩn của trường
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Quyết định trúng tuyển...")
    title = data.get("title", "Quyết định về việc công nhận thí sinh trúng tuyển")
    body = data.get("body", "") # Nội dung chính có thể chỉ là phần căn cứ, quyết định
    # Danh sách trúng tuyển thường nằm trong data['attachment'] hoặc data['student_list']
    student_list_info = data.get("student_list_info", "Danh sách kèm theo Quyết định này")
    issuing_org = data.get("issuing_org", "TÊN TRƯỜNG").upper()
    issuing_authority = data.get("issuing_authority", "HIỆU TRƯỞNG").upper() # Người ký QĐ

    # 1. Header trường
    data['issuing_org'] = issuing_org
    format_basic_header(document, data, "QuyetDinhTS")

    # 2. Tên loại
    p_tenloai = document.add_paragraph("QUYẾT ĐỊNH")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "QUYẾT ĐỊNH", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Trích yếu
    qd_title = title.replace("Quyết định", "").strip()
    p_title = document.add_paragraph(f"Về việc {qd_title}")
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, f"Về việc {qd_title}", size=Pt(14), bold=True)

    # 4. Thẩm quyền ban hành (Hiệu trưởng)
    p_authority = document.add_paragraph(issuing_authority)
    set_paragraph_format(p_authority, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_authority, issuing_authority, size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # 5. Căn cứ ban hành Quyết định
    preamble = data.get("preamble", [
        "Căn cứ Quy chế tổ chức và hoạt động của Trường...",
        "Căn cứ Quy chế tuyển sinh hiện hành của Bộ Giáo dục và Đào tạo;",
        "Căn cứ Đề án tuyển sinh của Trường năm ...;",
        "Căn cứ kết quả xét tuyển của Hội đồng tuyển sinh Trường ... ngày .../.../...;" ,
        "Xét đề nghị của Trưởng phòng Đào tạo,"
        ]) # Mẫu căn cứ
    if isinstance(preamble, list):
        for line in preamble:
            p_pre = document.add_paragraph(line)
            set_paragraph_format(p_pre, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, space_after=Pt(0), line_spacing=1.5)
            add_run_with_format(p_pre, line, size=FONT_SIZE_DEFAULT, italic=True)
    document.add_paragraph() # Khoảng trống

    # 6. QUYẾT ĐỊNH:
    p_qd_label = document.add_paragraph("QUYẾT ĐỊNH:")
    set_paragraph_format(p_qd_label, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
    add_run_with_format(p_qd_label, "QUYẾT ĐỊNH:", size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # 7. Nội dung Quyết định (Các Điều)
    # Điều 1: Công nhận danh sách...
    p_dieu1 = document.add_paragraph()
    set_paragraph_format(p_dieu1, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent=Cm(0))
    add_run_with_format(p_dieu1, "Điều 1.", bold=True)
    add_run_with_format(p_dieu1, f" Công nhận ... thí sinh trúng tuyển vào học chương trình ... năm ... của Trường {issuing_org} ")
    add_run_with_format(p_dieu1, f"({student_list_info}).") # Tham chiếu danh sách

    # Điều 2: Trách nhiệm thi hành
    p_dieu2 = document.add_paragraph()
    set_paragraph_format(p_dieu2, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent=Cm(0))
    add_run_with_format(p_dieu2, "Điều 2.", bold=True)
    add_run_with_format(p_dieu2, " Các Ông/Bà Trưởng phòng Đào tạo, Trưởng các đơn vị có liên quan và các thí sinh có tên tại Điều 1 chịu trách nhiệm thi hành Quyết định này.")

    # Điều 3: Hiệu lực thi hành
    p_dieu3 = document.add_paragraph()
    set_paragraph_format(p_dieu3, alignment=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent=Cm(0))
    add_run_with_format(p_dieu3, "Điều 3.", bold=True)
    add_run_with_format(p_dieu3, " Quyết định này có hiệu lực kể từ ngày ký.")

    # (Có thể thêm các Điều khác từ body nếu có)
    body_lines = body.split('\n')
    for line in body_lines:
         # Xử lý thêm các điều/nội dung khác nếu cần
         pass


    # 8. Chữ ký (Hiệu trưởng)
    if not data.get('signer_title'): data['signer_title'] = issuing_authority # Tự lấy chức danh Hiệu trưởng
    format_signature_block(document, data)

    # 9. Nơi nhận
    if not data.get('recipients'):
        data['recipients'] = [
            "- Như Điều 2;",
            "- Hội đồng tuyển sinh;",
            "- Lưu: VT, Phòng Đào tạo."
        ]
    format_recipient_list(document, data)

    # Ghi chú: Cần có cơ chế xử lý danh sách đính kèm riêng

    print("Định dạng Quyết định trúng tuyển hoàn tất.")
# formatters/quyet_dinh.py
import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
# Import các hàm tiện ích và thành phần chung
from utils import set_paragraph_format, set_run_format, add_run_with_format
from .common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    """Hàm chính định dạng cho loại văn bản Quyết Định."""
    print("Bắt đầu định dạng Quyết định...")
    title = data.get("title", "Quyết định về việc ABC")
    body = data.get("body", "Nội dung quyết định...") # Nội dung đã qua AI làm sạch
    authority_org = data.get("issuing_org", "CƠ QUAN BAN HÀNH").upper()

    # 1. Tạo Header chuẩn
    format_basic_header(document, data, "QuyetDinh")

    # 2. Tên loại và Trích yếu
    p_tenloai = document.add_paragraph("QUYẾT ĐỊNH")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    set_run_format(p_tenloai.runs[0], size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    trich_yeu_text = title.replace("Quyết định", "").strip()
    p_trichyeu = document.add_paragraph(f"Về việc {trich_yeu_text}")
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    set_run_format(p_trichyeu.runs[0], size=Pt(14), bold=True)
    # Gạch ngang dưới trích yếu
    p_line_ty = document.add_paragraph("-" * 15)
    set_paragraph_format(p_line_ty, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 3. Thẩm quyền ban hành
    p_thamquyen = document.add_paragraph(authority_org)
    set_paragraph_format(p_thamquyen, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    set_run_format(p_thamquyen.runs[0], size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # 4. Nội dung (Căn cứ, Quyết định, Điều)
    body_lines = body.split('\n')
    căn_cứ_ended = False
    body_content_started = False
    # (Copy logic xử lý Căn cứ, QUYẾT ĐỊNH:, Điều từ app.py cũ vào đây)
    # ... (Logic xử lý body cho Quyết định) ...
    for i, line in enumerate(body_lines):
        stripped_line = line.strip()
        if not stripped_line: continue

        if stripped_line.lower().startswith("căn cứ") and not căn_cứ_ended and not body_content_started:
             p = document.add_paragraph()
             set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(0))
             run = p.add_run(stripped_line)
             set_run_format(run, size=FONT_SIZE_DEFAULT, italic=True)
             if i + 1 >= len(body_lines) or not body_lines[i+1].strip().lower().startswith("căn cứ"):
                 căn_cứ_ended = True
                 if i + 1 < len(body_lines) and body_lines[i+1].strip().lower().startswith("xét đề nghị"):
                     # Xử lý "Xét đề nghị"
                     pass # (Thêm logic nếu cần)
                 # Thêm chữ QUYẾT ĐỊNH:
                 p_qd_label = document.add_paragraph("QUYẾT ĐỊNH:")
                 set_paragraph_format(p_qd_label, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
                 set_run_format(p_qd_label.runs[0], size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)
                 body_content_started = True
        elif căn_cứ_ended and body_content_started and stripped_line:
             p = document.add_paragraph()
             is_dieu = stripped_line.upper().startswith("ĐIỀU")
             align = WD_ALIGN_PARAGRAPH.LEFT if is_dieu else WD_ALIGN_PARAGRAPH.JUSTIFY
             indent = Cm(0) if is_dieu else FIRST_LINE_INDENT
             set_paragraph_format(p, alignment=align, first_line_indent=indent, line_spacing=1.5, space_after=Pt(6))
             run = p.add_run(stripped_line)
             set_run_format(run, size=FONT_SIZE_DEFAULT, bold=is_dieu)
        # ... (Các xử lý khác)


    # 5. Chữ ký
    format_signature_block(document, data)

    # 6. Nơi nhận
    format_recipient_list(document, data)

    print("Định dạng Quyết định hoàn tất.")

# (Tạo các file tương tự cho cong_van.py, chi_thi.py, thong_bao.py, ke_hoach.py và chuyển logic định dạng vào đó)
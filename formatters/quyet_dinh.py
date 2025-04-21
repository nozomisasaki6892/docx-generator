# formatters/quyet_dinh.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Quyết định...")
    title = data.get("title", "Quyết định về việc ABC")
    body = data.get("body", "Nội dung quyết định...")
    authority_org = data.get("issuing_org", "CƠ QUAN BAN HÀNH").upper()

    format_basic_header(document, data, "QuyetDinh")

    p_tenloai = document.add_paragraph("QUYẾT ĐỊNH")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "QUYẾT ĐỊNH", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    trich_yeu_text = title.replace("Quyết định", "").strip()
    p_trichyeu = document.add_paragraph(f"Về việc {trich_yeu_text}")
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_trichyeu, f"Về việc {trich_yeu_text}", size=Pt(14), bold=True)

    p_line_ty = document.add_paragraph("-" * 15)
    set_paragraph_format(p_line_ty, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    p_thamquyen = document.add_paragraph(authority_org)
    set_paragraph_format(p_thamquyen, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_thamquyen.runs[0], size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    body_lines = body.split('\n')
    căn_cứ_ended = False
    body_content_started = False
    processed_indices = set()
    has_quyet_dinh_label = False

    for i, line in enumerate(body_lines):
        if i in processed_indices: continue
        stripped_line = line.strip()
        if not stripped_line: continue

        if stripped_line.lower().startswith("căn cứ") and not căn_cứ_ended and not body_content_started:
            p = document.add_paragraph()
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(0))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, italic=True)
            processed_indices.add(i)
            if i + 1 >= len(body_lines) or not body_lines[i+1].strip().lower().startswith("căn cứ"):
                 căn_cứ_ended = True
                 next_line_index = i + 1
                 if next_line_index < len(body_lines) and body_lines[next_line_index].strip().lower().startswith("xét đề nghị"):
                     p_xet = document.add_paragraph()
                     set_paragraph_format(p_xet, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))
                     add_run_with_format(p_xet, body_lines[next_line_index].strip(), size=FONT_SIZE_DEFAULT, italic=True)
                     processed_indices.add(next_line_index)
                     next_line_index += 1 # Move past "Xét đề nghị"

                 # Check if body already contains QUYẾT ĐỊNH:
                 if next_line_index < len(body_lines) and "QUYẾT ĐỊNH:" in body_lines[next_line_index].upper():
                     has_quyet_dinh_label = True
                     # Process this line as regular content below
                 else:
                     p_qd_label = document.add_paragraph("QUYẾT ĐỊNH:")
                     set_paragraph_format(p_qd_label, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
                     add_run_with_format(p_qd_label, "QUYẾT ĐỊNH:", size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)
                 body_content_started = True


        elif (căn_cứ_ended or not any(l.strip().lower().startswith("căn cứ") for l in body_lines)) and stripped_line:
             # Xử lý phần nội dung sau căn cứ hoặc nếu không có căn cứ
             if not body_content_started: # Nếu chưa có QUYẾT ĐỊNH: thì thêm vào
                  if not has_quyet_dinh_label and "QUYẾT ĐỊNH:" not in stripped_line.upper():
                       p_qd_label = document.add_paragraph("QUYẾT ĐỊNH:")
                       set_paragraph_format(p_qd_label, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
                       add_run_with_format(p_qd_label, "QUYẾT ĐỊNH:", size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)
                  body_content_started = True

             p = document.add_paragraph()
             is_dieu = stripped_line.upper().startswith("ĐIỀU")
             # Điều: Căn trái, đậm
             # Khoản (1., 2.): Thụt lề trái, không thụt dòng đầu
             # Điểm (a), b)): Thụt lề sâu hơn
             is_khoan = re.match(r'^\d+\.\s+', stripped_line)
             is_diem = re.match(r'^[a-z]\)\s+', stripped_line)

             align = WD_ALIGN_PARAGRAPH.JUSTIFY
             left_indent = Cm(0)
             first_indent = FIRST_LINE_INDENT

             if is_dieu:
                 align = WD_ALIGN_PARAGRAPH.LEFT
                 first_indent = Cm(0)
             elif is_khoan:
                 left_indent = Cm(0.5)
                 first_indent = Cm(0)
             elif is_diem:
                 left_indent = Cm(1.0)
                 first_indent = Cm(0)

             set_paragraph_format(p, alignment=align, left_indent=left_indent, first_line_indent=first_indent, line_spacing=1.5, space_after=Pt(6))
             add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_dieu) # Chỉ in đậm Điều
             processed_indices.add(i)

    format_signature_block(document, data)
    format_recipient_list(document, data)
    print("Định dạng Quyết định hoàn tất.")
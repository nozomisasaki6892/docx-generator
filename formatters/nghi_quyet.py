# formatters/nghi_quyet.py
import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_signature_block, format_recipient_list
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT

def format(document, data):
    print("Bắt đầu định dạng Nghị quyết...")
    title = data.get("title", "Nghị quyết về việc ABC")
    body = data.get("body", "Nội dung nghị quyết...")
    # Nghị quyết thường do cơ quan tập thể ban hành (HĐND, Chính phủ,...)
    issuing_authority = data.get("issuing_org", "CƠ QUAN BAN HÀNH").upper() # VD: HỘI ĐỒNG NHÂN DÂN TỈNH XYZ

    # 1. Header (Thường căn giữa Tên cơ quan)
    format_basic_header(document, data, "NghiQuyet")

    # 2. Tên loại
    p_tenloai = document.add_paragraph("NGHỊ QUYẾT")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "NGHỊ QUYẾT", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # 3. Trích yếu
    trich_yeu_text = title.replace("Nghị quyết", "").strip()
    p_trichyeu = document.add_paragraph(f"Về việc {trich_yeu_text}")
    set_paragraph_format(p_trichyeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_trichyeu, f"Về việc {trich_yeu_text}", size=Pt(14), bold=True)
    p_line_ty = document.add_paragraph("-" * 15)
    set_paragraph_format(p_line_ty, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))

    # 4. Cơ quan ban hành (lặp lại dưới dạng text)
    p_issuer = document.add_paragraph(issuing_authority)
    set_paragraph_format(p_issuer, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
    add_run_with_format(p_issuer, issuing_authority, size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)

    # 5. Nội dung (Căn cứ, Phần QUYẾT NGHỊ:, các Điều)
    body_lines = body.split('\n')
    can_cu_ended = False
    body_content_started = False
    processed_indices = set()
    has_quyet_nghi_label = False

    for i, line in enumerate(body_lines):
        if i in processed_indices: continue
        stripped_line = line.strip()
        if not stripped_line: continue

        # Xử lý Căn cứ
        if stripped_line.lower().startswith("căn cứ") and not can_cu_ended and not body_content_started:
            p = document.add_paragraph()
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(0))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, italic=True)
            processed_indices.add(i)
            # Check if this is the last Căn cứ line
            if i + 1 >= len(body_lines) or not body_lines[i+1].strip().lower().startswith("căn cứ"):
                can_cu_ended = True
                # Check for QUYẾT NGHỊ: label next
                next_line_idx = i + 1
                if next_line_idx < len(body_lines) and "QUYẾT NGHỊ:" in body_lines[next_line_idx].upper():
                    has_quyet_nghi_label = True
                else: # Add label if missing
                    p_qn_label = document.add_paragraph("QUYẾT NGHỊ:")
                    set_paragraph_format(p_qn_label, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
                    add_run_with_format(p_qn_label, "QUYẾT NGHỊ:", size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)
                body_content_started = True

        # Xử lý nội dung sau căn cứ
        elif (can_cu_ended or not any(l.strip().lower().startswith("căn cứ") for l in body_lines)) and stripped_line:
            if not body_content_started: # Nếu chưa có QUYẾT NGHỊ:
                 if not has_quyet_nghi_label and "QUYẾT NGHỊ:" not in stripped_line.upper():
                      p_qn_label = document.add_paragraph("QUYẾT NGHỊ:")
                      set_paragraph_format(p_qn_label, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6))
                      add_run_with_format(p_qn_label, "QUYẾT NGHỊ:", size=FONT_SIZE_DEFAULT, bold=True, uppercase=True)
                 body_content_started = True

            p = document.add_paragraph()
            is_dieu = stripped_line.upper().startswith("ĐIỀU")
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
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_dieu)
            processed_indices.add(i)

    # 6. Chữ ký (Thường là TM. Cơ quan tập thể, CHỦ TỊCH/TRƯỞNG BAN...)
    # Cần truyền đúng authority_signer và signer_title từ data
    if not data.get('authority_signer'): data['authority_signer'] = f"TM. {issuing_authority}"
    if not data.get('signer_title'): data['signer_title'] = "CHỦ TỊCH" # Hoặc chức vụ phù hợp
    format_signature_block(document, data)

    # 7. Nơi nhận
    format_recipient_list(document, data)

    print("Định dạng Nghị quyết hoàn tất.")
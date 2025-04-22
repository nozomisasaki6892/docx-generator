import re
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import set_paragraph_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_recipient_list

from config import FONT_SIZE_DEFAULT, FONT_SIZE_SMALL, FIRST_LINE_INDENT, FONT_SIZE_SIGNATURE, FONT_SIZE_SIGNER_NAME, FONT_SIZE_VV

def format(document, data):
    print("Đang định dạng Công văn...")
    title = data.get("title", "Về việc ABC")
    body = data.get("body", "")
    recipients_to = data.get("recipients_to", "Kính gửi: [Tên đơn vị/cá nhân]")
    signer_title = data.get("signer_title", "CHỨC VỤ NGƯỜI KÝ").upper()
    signer_name = data.get("signer_name", "Người Ký")

    # 1. Tiêu ngữ + CQBH
    format_basic_header(document, data, "CongVan")

    # 2. Trích yếu (V/v:)
    trich_yeu_text = title
    if title.lower().startswith("về việc"):
        trich_yeu_text = title[len("về việc"):].strip()
    elif title.lower().startswith("v/v"):
        trich_yeu_text = title[len("v/v"):].strip()

    p_vv = document.add_paragraph()
    set_paragraph_format(p_vv, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(0), space_after=Pt(6))
    add_run_with_format(p_vv, f"V/v: {trich_yeu_text}", size=FONT_SIZE_VV)

    # 3. Kính gửi
    p_kg = document.add_paragraph()
    set_paragraph_format(p_kg, alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=FIRST_LINE_INDENT, space_before=Pt(6), space_after=Pt(6))
    add_run_with_format(p_kg, recipients_to, size=FONT_SIZE_DEFAULT, bold=True)

    # 4. Nội dung
    body_lines = body.split('\n')
    for line in body_lines:
        line = line.strip()
        if not line:
            continue
        p = document.add_paragraph()
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent=FIRST_LINE_INDENT, line_spacing=1.5, space_after=Pt(6))
        add_run_with_format(p, line, size=FONT_SIZE_DEFAULT)

    # 5. Chữ ký (ưu tiên dữ liệu từ AI)
    p_sig = document.add_paragraph()
    set_paragraph_format(p_sig, alignment=WD_ALIGN_PARAGRAPH.RIGHT, space_before=Pt(12), space_after=Pt(0), line_spacing=1.0)
    add_run_with_format(p_sig, signer_title + "\n\n\n\n\n", size=FONT_SIZE_SIGNATURE, bold=True)
    add_run_with_format(p_sig, signer_name, size=FONT_SIZE_SIGNER_NAME, bold=True)

    # 6. Nơi nhận
    format_recipient_list(document, data)

    print("Định dạng Công văn hoàn tất.")

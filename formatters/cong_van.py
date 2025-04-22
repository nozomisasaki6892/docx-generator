import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

from utils import set_paragraph_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_recipient_list
except ImportError:
    from common_elements import format_basic_header, format_recipient_list

from config import (
    FONT_NAME, FONT_SIZE_DEFAULT, FONT_SIZE_SMALL, FIRST_LINE_INDENT,
    FONT_SIZE_VV, FONT_SIZE_SIGNATURE, FONT_SIZE_SIGNER_NAME
)

def format(document, data):
    """Định dạng tài liệu Word theo mẫu Công văn chuẩn."""
    print("Đang định dạng Công văn...")
    title = data.get("title", "Về việc [Nội dung trích yếu]")
    body = data.get("body", "[Nội dung chính của công văn]")
    recipients_to = data.get("recipients_to", "Kính gửi: [Tên đơn vị/cá nhân nhận]")

    format_basic_header(document, data, "CongVan")

    trich_yeu_text = title
    if title.lower().startswith("về việc"):
        trich_yeu_text = title[len("về việc"):].strip()
    elif title.lower().startswith("v/v"):
        trich_yeu_text = title[len("v/v"):].strip()

    p_vv = document.add_paragraph()
    set_paragraph_format(
        p_vv,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=Pt(6),
        space_after=Pt(6)
    )
    add_run_with_format(
        p_vv,
        f"V/v: {trich_yeu_text}",
        size=FONT_SIZE_VV,
        bold=True
    )

    p_kg = document.add_paragraph()
    set_paragraph_format(
        p_kg,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=FIRST_LINE_INDENT,
        space_before=Pt(6),
        space_after=Pt(6)
    )
    add_run_with_format(p_kg, recipients_to, size=FONT_SIZE_DEFAULT, bold=True)

    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            set_paragraph_format(
                p,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                first_line_indent=FIRST_LINE_INDENT,
                line_spacing=1.5,
                space_after=Pt(6)
            )
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT)

    signer_title = data.get("signer_title", "").upper()
    signer_name = data.get("signer_name", "")

    if signer_title or signer_name:
        p_sig_title = document.add_paragraph()
        set_paragraph_format(
            p_sig_title,
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            space_before=Pt(12),
            space_after=Pt(0),
            line_spacing=1.0
        )
        add_run_with_format(
            p_sig_title,
            signer_title if signer_title else "[CHỨC VỤ NGƯỜI KÝ]",
            size=FONT_SIZE_SIGNATURE,
            bold=True
        )

        p_sig_space = document.add_paragraph()
        set_paragraph_format(
            p_sig_space,
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            space_before=Pt(0),
            space_after=Pt(0),
            line_spacing=1.0
        )
        add_run_with_format(p_sig_space, "\n\n\n", size=FONT_SIZE_SIGNER_NAME)

        p_sig_name = document.add_paragraph()
        set_paragraph_format(
            p_sig_name,
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
            space_before=Pt(0),
            space_after=Pt(0),
            line_spacing=1.0
        )
        add_run_with_format(
            p_sig_name,
            signer_name if signer_name else "[Họ và tên người ký]",
            size=FONT_SIZE_SIGNER_NAME,
            bold=True
        )

    format_recipient_list(document, data)

    print("Định dạng Công văn hoàn tất.")
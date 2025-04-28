# utils.py
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from config import FONT_NAME # Giả sử config.py chỉ chứa FONT_NAME

def set_paragraph_format(
    paragraph,
    alignment=None,
    left_indent=None,
    right_indent=None,
    first_line_indent=None,
    space_before=None,
    space_after=None,
    line_spacing=None,
    line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    keep_together=None,
    keep_with_next=None,
    page_break_before=None,
    widow_control=True
):
    """Thiết lập định dạng cho đối tượng Paragraph."""
    paragraph_format = paragraph.paragraph_format
    if alignment is not None:
        paragraph_format.alignment = alignment
    if left_indent is not None:
        paragraph_format.left_indent = left_indent
    if right_indent is not None:
        paragraph_format.right_indent = right_indent
    if first_line_indent is not None:
        paragraph_format.first_line_indent = first_line_indent
    if space_before is not None:
        paragraph_format.space_before = space_before
    if space_after is not None:
        paragraph_format.space_after = space_after
    if line_spacing is not None:
        paragraph_format.line_spacing = line_spacing
        paragraph_format.line_spacing_rule = line_spacing_rule
    if keep_together is not None:
        paragraph_format.keep_together = keep_together
    if keep_with_next is not None:
        paragraph_format.keep_with_next = keep_with_next
    if page_break_before is not None:
        paragraph_format.page_break_before = page_break_before
    if widow_control is not None:
        paragraph_format.widow_control = widow_control

def set_run_format(
    run,
    font_name=FONT_NAME,
    size=None,
    bold=None,
    italic=None,
    underline=None,
    uppercase=None
):
    """Thiết lập định dạng cho đối tượng Run."""
    font = run.font
    font.name = font_name
    # Đảm bảo font chữ áp dụng cho ký tự Đông Á (quan trọng cho tiếng Việt)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    if size is not None:
        font.size = size
    if bold is not None:
        font.bold = bold
    if italic is not None:
        font.italic = italic
    if underline is not None:
        font.underline = underline
    if uppercase is not None:
        run.text = run.text.upper()

def add_run_with_format(
    paragraph,
    text,
    font_name=FONT_NAME,
    size=None,
    bold=None,
    italic=None,
    underline=None,
    uppercase=None
):
    """Thêm một Run vào Paragraph và áp dụng định dạng."""
    run = paragraph.add_run(text)
    set_run_format(run, font_name, size, bold, italic, underline, uppercase)
    return run

def add_paragraph_with_text(
    document,
    text,
    alignment=None,
    left_indent=None,
    right_indent=None,
    first_line_indent=None,
    space_before=None,
    space_after=None,
    line_spacing=None,
    line_spacing_rule=WD_LINE_SPACING.MULTIPLE,
    keep_together=None,
    keep_with_next=None,
    page_break_before=None,
    widow_control=True,
    font_name=FONT_NAME,
    size=None,
    bold=None,
    italic=None,
    underline=None,
    uppercase=None
):
    """Thêm một Paragraph với text và định dạng đầy đủ."""
    paragraph = document.add_paragraph()
    set_paragraph_format(
        paragraph, alignment, left_indent, right_indent, first_line_indent,
        space_before, space_after, line_spacing, line_spacing_rule,
        keep_together, keep_with_next, page_break_before, widow_control
    )
    add_run_with_format(
        paragraph, text, font_name, size, bold, italic, underline, uppercase
    )
    return paragraph

def apply_standard_margins(document):
    """Áp dụng lề trang chuẩn theo Nghị định 30."""
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)    # Tối thiểu 20mm
    section.bottom_margin = Cm(2.0) # Tối thiểu 20mm
    section.left_margin = Cm(3.0)   # Tối thiểu 30mm
    section.right_margin = Cm(1.5)  # Tối thiểu 15mm

def apply_contract_margins(document):
    """Áp dụng lề trang cho Hợp đồng (ví dụ)."""
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(3.0) # Giữ nguyên lề trái rộng
    section.right_margin = Cm(1.5) # Có thể điều chỉnh nếu cần

def apply_landscape_margins(document, top=1.5, bottom=1.5, left=1.5, right=1.5):
    """Áp dụng lề trang cho trang ngang (ví dụ Bằng tốt nghiệp)."""
    section = document.sections[0]
    section.orientation = 1 # WD_ORIENTATION.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)

def add_horizontal_line(paragraph, length_chars=20):
    """Thêm một dòng kẻ ngang đơn giản."""
    # Cách này đơn giản nhưng không đẹp bằng shape
    run = paragraph.add_run('_' * length_chars)
    set_run_format(run, bold=True) # Theo NĐ30, dòng kẻ đậm
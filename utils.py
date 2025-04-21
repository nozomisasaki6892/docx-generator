# utils.py
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from config import FONT_NAME, FONT_SIZE_DEFAULT, FIRST_LINE_INDENT, LINE_SPACING_DEFAULT

def set_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0), first_line_indent=Cm(0), space_before=Pt(0), space_after=Pt(6), line_spacing=LINE_SPACING_DEFAULT, line_spacing_rule=WD_LINE_SPACING.MULTIPLE, keep_together=False, keep_with_next=False, page_break_before=False):
    """Thiết lập định dạng chung cho một Paragraph."""
    p_format = paragraph.paragraph_format
    p_format.alignment = alignment
    p_format.left_indent = left_indent
    p_format.right_indent = Cm(0)
    p_format.first_line_indent = first_line_indent
    p_format.space_before = space_before
    p_format.space_after = space_after
    p_format.line_spacing = line_spacing
    p_format.line_spacing_rule = line_spacing_rule
    p_format.keep_together = keep_together
    p_format.keep_with_next = keep_with_next
    p_format.page_break_before = page_break_before

def set_run_format(run, font_name=FONT_NAME, size=FONT_SIZE_DEFAULT, bold=False, italic=False, underline=False, uppercase=False):
    """Thiết lập định dạng cho một Run (phần text)."""
    font = run.font
    font.name = font_name
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    except Exception as e:
        print(f"Warning: Không thể set eastAsia font - {e}")
    font.size = size
    font.bold = bold
    font.italic = italic
    font.underline = underline
    run.text = run.text.upper() if uppercase else run.text

def add_run_with_format(paragraph, text, size=FONT_SIZE_DEFAULT, bold=False, italic=False, uppercase=False, font_name=FONT_NAME):
    """Thêm một Run vào Paragraph và định dạng nó."""
    run = paragraph.add_run(text)
    set_run_format(run, font_name=font_name, size=size, bold=bold, italic=italic, uppercase=uppercase)
    return run

# Thêm các hàm tiện ích khác nếu cần (ví dụ: tạo bảng, chèn ảnh...)
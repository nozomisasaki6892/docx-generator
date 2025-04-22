# utils.py (Phiên bản chuẩn hóa cuối cùng)
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn # Import qn để set East Asia font
# Import các config cần thiết làm giá trị mặc định
from config import FONT_NAME, FONT_SIZE_DEFAULT, FIRST_LINE_INDENT, LINE_SPACING_DEFAULT

def set_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, # Mặc định căn đều
                         left_indent=None, right_indent=None,
                         first_line_indent=FIRST_LINE_INDENT, # Mặc định thụt lề đầu dòng từ config
                         space_before=Pt(0), space_after=Pt(6), # Mặc định khoảng cách sau là 6pt
                         line_spacing=LINE_SPACING_DEFAULT, # Mặc định giãn dòng từ config
                         line_spacing_rule=WD_LINE_SPACING.MULTIPLE, # Quy tắc giãn dòng mặc định
                         keep_together=False, keep_with_next=False, page_break_before=False,
                         widow_control=True):
    """Thiết lập định dạng chung cho một Paragraph, ưu tiên các giá trị truyền vào."""
    p_format = paragraph.paragraph_format
    p_format.alignment = alignment
    # Chỉ đặt nếu giá trị được truyền vào khác None, nếu không giữ mặc định của Word/Style
    if left_indent is not None:
        p_format.left_indent = left_indent
    if right_indent is not None:
        p_format.right_indent = right_indent
    # Thụt lề đầu dòng có giá trị mặc định từ config
    p_format.first_line_indent = first_line_indent
    # Khoảng cách có giá trị mặc định cụ thể
    p_format.space_before = space_before
    p_format.space_after = space_after
    # Giãn dòng có giá trị mặc định từ config
    p_format.line_spacing = line_spacing
    p_format.line_spacing_rule = line_spacing_rule

    p_format.keep_together = keep_together
    p_format.keep_with_next = keep_with_next
    p_format.page_break_before = page_break_before
    p_format.widow_control = widow_control


def set_run_format(run, font_name=FONT_NAME, size=None, bold=False, italic=False,
                   underline=False, color_rgb=None, uppercase=False, subscript=False, superscript=False):
    """Thiết lập định dạng cho một Run (phần text)."""
    r_font = run.font
    r_font.name = font_name
    # --- Thêm lại phần set East Asia font ---
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    except Exception as e:
        print(f"Warning: Không thể set eastAsia font - {e}")
    # --- Kết thúc phần thêm lại ---

    # Set size, ưu tiên giá trị truyền vào, nếu không thì dùng FONT_SIZE_DEFAULT
    effective_size = size if size is not None else FONT_SIZE_DEFAULT
    if isinstance(effective_size, (int, float)):
        r_font.size = Pt(effective_size)
    else:
        r_font.size = effective_size # Giả sử đã là Pt object

    run.bold = bold
    run.italic = italic
    run.underline = underline
    if color_rgb is not None:
        from docx.shared import RGBColor
        r_font.color.rgb = RGBColor.from_string(color_rgb) # Expecting "FF0000" format
    if uppercase:
        run.text = run.text.upper()

    run.subscript = subscript
    run.superscript = superscript


def add_run_with_format(paragraph, text, font_name=FONT_NAME, size=None,
                        bold=False, italic=False, underline=False, color_rgb=None,
                        uppercase=False, subscript=False, superscript=False):
    """Thêm một Run vào Paragraph và định dạng nó."""
    run = paragraph.add_run(text.upper() if uppercase else text)
    # Gọi hàm set_run_format đã cập nhật
    set_run_format(run, font_name=font_name, size=size, bold=bold, italic=italic,
                   underline=underline, color_rgb=color_rgb, uppercase=False, # uppercase xử lý khi add_run
                   subscript=subscript, superscript=superscript)
    return run


def add_centered_text(document, text, font_name=FONT_NAME, size=None,
                      bold=False, italic=False, underline=False, color_rgb=None,
                      uppercase=False, space_before=None, space_after=None, line_spacing=1.0):
    """Thêm một đoạn văn bản căn giữa với định dạng chỉ định."""
    p = document.add_paragraph()
    # Dùng hàm set_paragraph_format đã cập nhật, chỉ định alignment và spacing
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=Cm(0), # Căn giữa không thụt lề
                         space_before=space_before if space_before is not None else Pt(0),
                         space_after=space_after if space_after is not None else Pt(6), # Mặc định space after 6
                         line_spacing=line_spacing)
    add_run_with_format(p, text, font_name=font_name, size=size, bold=bold, italic=italic,
                        underline=underline, color_rgb=color_rgb, uppercase=uppercase)
    return p
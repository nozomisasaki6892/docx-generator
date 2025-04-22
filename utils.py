# utils.py (Dựa trên bản gốc + bổ sung add_centered_text)
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn # Giữ lại import quan trọng này
from config import FONT_NAME, FONT_SIZE_DEFAULT, FIRST_LINE_INDENT, LINE_SPACING_DEFAULT

def set_paragraph_format(paragraph, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, left_indent=Cm(0), first_line_indent=FIRST_LINE_INDENT, # Dùng FIRST_LINE_INDENT từ config làm mặc định
                         space_before=Pt(0), space_after=Pt(6), line_spacing=LINE_SPACING_DEFAULT, # Dùng LINE_SPACING_DEFAULT từ config
                         line_spacing_rule=WD_LINE_SPACING.MULTIPLE, keep_together=False, keep_with_next=False, page_break_before=False, widow_control=True): # Thêm widow_control
    """Thiết lập định dạng chung cho một Paragraph."""
    p_format = paragraph.paragraph_format
    p_format.alignment = alignment
    p_format.left_indent = left_indent
    p_format.right_indent = Cm(0) # Giữ lại thiết lập right_indent=0
    p_format.first_line_indent = first_line_indent
    p_format.space_before = space_before
    p_format.space_after = space_after
    p_format.line_spacing = line_spacing
    p_format.line_spacing_rule = line_spacing_rule
    p_format.keep_together = keep_together
    p_format.keep_with_next = keep_with_next
    p_format.page_break_before = page_break_before
    p_format.widow_control = widow_control # Thêm widow_control

def set_run_format(run, font_name=FONT_NAME, size=FONT_SIZE_DEFAULT, bold=False, italic=False, underline=False, uppercase=False, color_rgb=None, subscript=False, superscript=False): # Thêm các tùy chọn định dạng
    """Thiết lập định dạng cho một Run (phần text)."""
    font = run.font
    font.name = font_name
    try:
        # Giữ lại thiết lập East Asia font quan trọng
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    except Exception as e:
        print(f"Warning: Không thể set eastAsia font - {e}")

    if isinstance(size, (int, float)): # Cho phép truyền số vào size
         font.size = Pt(size)
    else:
         font.size = size # Mặc định FONT_SIZE_DEFAULT

    font.bold = bold
    font.italic = italic
    font.underline = underline # Thêm underline

    if color_rgb is not None: # Thêm màu sắc
        from docx.shared import RGBColor
        font.color.rgb = RGBColor.from_string(color_rgb)

    run.subscript = subscript # Thêm chỉ số dưới
    run.superscript = superscript # Thêm chỉ số trên

    # Xử lý uppercase sau cùng để tránh ghi đè text nếu các thuộc tính khác được set sau
    if uppercase:
         run.text = run.text.upper()


def add_run_with_format(paragraph, text, font_name=FONT_NAME, size=FONT_SIZE_DEFAULT, bold=False, italic=False, underline=False, uppercase=False, color_rgb=None, subscript=False, superscript=False): # Thêm các tham số
    """Thêm một Run vào Paragraph và định dạng nó."""
    # Xử lý uppercase trước khi add_run
    run_text = text.upper() if uppercase else text
    run = paragraph.add_run(run_text)
    # Gọi hàm set_run_format đầy đủ
    set_run_format(run, font_name=font_name, size=size, bold=bold, italic=italic,
                   underline=underline, color_rgb=color_rgb, uppercase=False, # Đã xử lý uppercase
                   subscript=subscript, superscript=superscript)
    return run

# --- Hàm add_centered_text được bổ sung ---
def add_centered_text(document, text, font_name=FONT_NAME, size=None,
                      bold=False, italic=False, underline=False, color_rgb=None,
                      uppercase=False, space_before=None, space_after=None, line_spacing=1.0):
    """Thêm một đoạn văn bản căn giữa với định dạng chỉ định."""
    p = document.add_paragraph()
    effective_space_before = space_before if space_before is not None else Pt(0)
    effective_space_after = space_after if space_after is not None else Pt(6) # Mặc định space after 6
    effective_size = size if size is not None else FONT_SIZE_DEFAULT

    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=Cm(0), # Căn giữa không thụt lề
                         space_before=effective_space_before,
                         space_after=effective_space_after,
                         line_spacing=line_spacing)
    add_run_with_format(p, text, font_name=font_name, size=effective_size, bold=bold, italic=italic,
                        underline=underline, color_rgb=color_rgb, uppercase=uppercase)
    return p
# formatters/bien_ban.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import set_paragraph_format, set_run_format, add_run_with_format
try:
    from .common_elements import format_basic_header, format_signature_block # Biên bản thường ko có Nơi nhận
except ImportError:
    from common_elements import format_basic_header, format_signature_block
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FIRST_LINE_INDENT, FONT_SIZE_HEADER

def format(document, data):
    print("Bắt đầu định dạng Biên bản...")
    title = data.get("title", "Biên bản họp ABC")
    body = data.get("body", "Nội dung biên bản...")
    # Biên bản có thể có hoặc không có header chuẩn, tùy thuộc vào ngữ cảnh
    # Nếu là biên bản của cuộc họp nội bộ thì không cần header NĐ30
    # Nếu là biên bản làm việc giữa các cơ quan thì có thể có
    add_header = data.get("add_formal_header", False) # Thêm cờ để quyết định có header ko

    if add_header:
        format_basic_header(document, data, "BienBan") # Có thể dùng header chung

    # Quốc hiệu/Tiêu ngữ (Thường có ở các biên bản quan trọng)
    p_qh = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    set_paragraph_format(p_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_qh, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", size=FONT_SIZE_HEADER, bold=True)
    p_tn = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    set_paragraph_format(p_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_tn, "Độc lập - Tự do - Hạnh phúc", size=Pt(13), bold=True)
    p_line_tn = document.add_paragraph("-" * 20)
    set_paragraph_format(p_line_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))


    # Tên loại
    p_tenloai = document.add_paragraph("BIÊN BẢN")
    set_paragraph_format(p_tenloai, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_tenloai, "BIÊN BẢN", size=FONT_SIZE_TITLE, bold=True, uppercase=True)

    # Tiêu đề của Biên bản
    bb_title = title.replace("Biên bản", "").strip()
    p_title = document.add_paragraph(bb_title)
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
    add_run_with_format(p_title, bb_title, size=Pt(14), bold=True)

    # Nội dung (Thời gian, Địa điểm, Thành phần, Chủ trì, Thư ký, Nội dung họp, Kết luận...)
    # Phần này rất đa dạng, cần AI tách hoặc có cấu trúc đầu vào rõ ràng
    # Tạm thời định dạng các dòng theo kiểu thông thường
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            # Các đề mục như Thời gian, Địa điểm, Thành phần,... có thể in đậm
            is_heading = any(stripped_line.upper().startswith(h) for h in ["THỜI GIAN:", "ĐỊA ĐIỂM:", "THÀNH PHẦN:", "CHỦ TRÌ:", "THƯ KÝ:", "NỘI DUNG:", "KẾT LUẬN:", "BIỂU QUYẾT:"])
            align = WD_ALIGN_PARAGRAPH.LEFT if is_heading else WD_ALIGN_PARAGRAPH.JUSTIFY
            indent = Cm(0) if is_heading else FIRST_LINE_INDENT

            set_paragraph_format(p, alignment=align, left_indent=Cm(0), first_line_indent=indent, line_spacing=1.5, space_after=Pt(6))
            add_run_with_format(p, stripped_line, size=FONT_SIZE_DEFAULT, bold=is_heading)

    # Chữ ký (Thường có nhiều chữ ký: Thư ký, Chủ trì, Đại diện các bên...)
    # Cần cấu trúc data['signatures'] phức tạp hơn [{title: 'THƯ KÝ', name: 'A'}, {title: 'CHỦ TRÌ', name: 'B'}]
    # Tạm thời dùng chữ ký đơn
    # Thêm dòng "Biên bản kết thúc vào lúc..."
    p_end = document.add_paragraph(f"Biên bản kết thúc vào lúc ... giờ ... ngày ... tháng ... năm ...")
    set_paragraph_format(p_end, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(12))
    add_run_with_format(p_end, p_end.text, size=FONT_SIZE_DEFAULT, italic=True)
    document.add_paragraph() # Khoảng trống

    # Ví dụ chữ ký Chủ trì và Thư ký (cần điều chỉnh layout table hoặc tabstop)
    sig_table = document.add_table(rows=1, cols=2)
    sig_table.autofit = False
    sig_table.columns[0].width = Inches(3.0)
    sig_table.columns[1].width = Inches(3.0)

    cell_left = sig_table.cell(0, 0)
    cell_left._element.clear_content()
    p_left_title = cell_left.add_paragraph("THƯ KÝ") # Ví dụ
    set_paragraph_format(p_left_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_left_title, "THƯ KÝ", size=FONT_SIZE_SIGNATURE, bold=True)
    p_left_note = cell_left.add_paragraph("(Ký, ghi rõ họ tên)")
    set_paragraph_format(p_left_note, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60)) # Khoảng trống ký
    add_run_with_format(p_left_note, "(Ký, ghi rõ họ tên)", size=Pt(11), italic=True)
    # Tên thư ký (nếu có)
    # p_left_name = cell_left.add_paragraph(data.get("secretary_name", " "))
    # set_paragraph_format(p_left_name, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    # add_run_with_format(p_left_name, data.get("secretary_name", " "), size=FONT_SIZE_SIGNER_NAME, bold=True)

    cell_right = sig_table.cell(0, 1)
    cell_right._element.clear_content()
    p_right_title = cell_right.add_paragraph(data.get("signer_title", "CHỦ TRÌ").upper()) # Ví dụ
    set_paragraph_format(p_right_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(p_right_title, data.get("signer_title", "CHỦ TRÌ").upper(), size=FONT_SIZE_SIGNATURE, bold=True)
    p_right_note = cell_right.add_paragraph("(Ký, ghi rõ họ tên)")
    set_paragraph_format(p_right_note, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(60)) # Khoảng trống ký
    add_run_with_format(p_right_note, "(Ký, ghi rõ họ tên)", size=Pt(11), italic=True)
    p_right_name = cell_right.add_paragraph(data.get("signer_name", " "))
    set_paragraph_format(p_right_name, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_run_with_format(p_right_name, data.get("signer_name", " "), size=FONT_SIZE_SIGNER_NAME, bold=True)


    # Biên bản thường không có Nơi nhận trừ khi cần gửi đi
    # format_recipient_list(document, data)

    print("Định dạng Biên bản hoàn tất.")
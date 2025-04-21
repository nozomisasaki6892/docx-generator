# formatters/phat_bieu.py
import re
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_DEFAULT, FONT_SIZE_TITLE, FONT_NAME

def format(document, data):
    print("Bắt đầu định dạng Phát biểu...")
    title = data.get("title", "BÀI PHÁT BIỂU TẠI SỰ KIỆN ABC")
    body = data.get("body", "Kính thưa quý vị...")
    event_info = data.get("event_info", {"name": "[Tên sự kiện]", "location": "[Địa điểm]", "date": "[Ngày tháng]"})
    is_draft = data.get("is_draft", False) # Cờ báo có phải bản nháp không

    # 1. Chỉ dẫn Dự thảo (Nếu có)
    if is_draft:
        p_draft = document.add_paragraph("(DỰ THẢO)")
        # Có thể căn trái hoặc giữa tùy ý
        set_paragraph_format(p_draft, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6))
        add_run_with_format(p_draft, "(DỰ THẢO)", size=FONT_SIZE_DEFAULT, bold=True)

    # 2. Tiêu đề bài phát biểu
    p_speech_title = document.add_paragraph(title.upper())
    set_paragraph_format(p_speech_title, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(12), space_after=Pt(6))
    add_run_with_format(p_speech_title, title.upper(), size=Pt(16), bold=True) # Cỡ chữ to

    # 3. Thông tin sự kiện
    event_text = f"{event_info['name']}\n{event_info['location']}, {event_info['date']}"
    p_event = document.add_paragraph(event_text)
    set_paragraph_format(p_event, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(18))
    add_run_with_format(p_event, event_text, size=Pt(12), italic=True) # Nhỏ hơn, nghiêng

    # 4. Nội dung phát biểu
    body_lines = body.split('\n')
    for line in body_lines:
        stripped_line = line.strip()
        if stripped_line:
            p = document.add_paragraph()
            # Phát biểu thường căn trái hoặc đều, giãn dòng lớn
            set_paragraph_format(
                p,
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, # Hoặc LEFT
                first_line_indent=Cm(1.0), # Có thể thụt lề hoặc không
                line_spacing=1.5, # Hoặc WD_LINE_SPACING.DOUBLE
                space_before=Pt(6),
                space_after=Pt(12) # Khoảng cách đoạn lớn
            )
            # Có thể tìm và in đậm các đoạn nhấn mạnh (khó tự động)
            add_run_with_format(p, stripped_line, size=Pt(14)) # Cỡ chữ đọc 14pt

    # 5. Lời kết
    # Thường nằm trong body, ví dụ: "Xin trân trọng cảm ơn!"

    # Không có chữ ký, nơi nhận

    print("Định dạng Phát biểu hoàn tất.")
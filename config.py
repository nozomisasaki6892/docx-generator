# config.py
import os
from dotenv import load_dotenv
from docx.shared import Pt, Cm

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_API_KEY_HERE_IF_NO_ENV")
GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"

FONT_NAME = 'Times New Roman'
FONT_SIZE_DEFAULT = Pt(13)
FONT_SIZE_LARGE = Pt(14)
FONT_SIZE_SMALL = Pt(11)
FONT_SIZE_TITLE = Pt(14)
FONT_SIZE_HEADER = Pt(12)
FONT_SIZE_PLACE_DATE = Pt(13)
FONT_SIZE_SIGNATURE = Pt(13)
FONT_SIZE_SIGNER_NAME = Pt(14)
FONT_SIZE_RECIPIENT_LABEL = Pt(12)
FONT_SIZE_DOC_NUMBER = Pt(13)
FONT_SIZE_VV = Pt(12)  # Cỡ chữ V/v của Công văn

MARGIN_TOP = Cm(2.0)
MARGIN_BOTTOM = Cm(2.0)
MARGIN_LEFT_DEFAULT = Cm(3.0)
MARGIN_RIGHT_DEFAULT = Cm(1.5)
MARGIN_LEFT_CONTRACT = Cm(3.0)

FIRST_LINE_INDENT = Cm(1.0)
LINE_SPACING_DEFAULT = 1.5

AI_PROMPT_TEMPLATE = """
Bạn là một trợ lý biên tập viên tiếng Việt chuyên nghiệp. Hãy đọc kỹ nội dung dưới đây và thực hiện các việc sau:
1. Sửa lỗi chính tả, ngữ pháp.
2. Loại bỏ từ ngữ thừa, câu lặp, diễn đạt khó hiểu.
3. Đảm bảo văn phong mạch lạc, rõ ràng, trang trọng, phù hợp với ngữ cảnh văn bản hành chính/công việc.
4. Giữ nguyên ý nghĩa gốc và các thông tin quan trọng như tên riêng, số liệu, địa danh.
5. Nếu phát hiện các phần tiêu đề không chuẩn như “DỰ ÁN”, “CÔNG VĂN”, “VĂN BẢN” hoặc các cụm từ không cần thiết ở đầu văn bản — hãy **xóa toàn bộ phần tiêu đề đó**.
6. Nếu nội dung có chứa phần chữ ký hoặc lời kết như:
   - "Trân trọng", "Kính thư", "TM.", "KT.", "Ký tên", "(Ký tên, đóng dấu)", "[Chức vụ]", "[Tên người ký]", tên thật của người ký, v.v...
   - Hãy **trích xuất tên người ký và chức danh**, và đặt vào 2 thẻ riêng như sau:
     [SIGNATURE_POSITION]Chức danh[/SIGNATURE_POSITION]
     [SIGNATURE_NAME]Tên người ký[/SIGNATURE_NAME]
   - Sau đó, **xóa toàn bộ phần chữ ký này khỏi nội dung chính** để tránh bị trùng chữ ký.
7. KHÔNG thêm: Quốc hiệu, Tiêu ngữ, Số ký hiệu, Ngày tháng, Nơi nhận, Chữ ký. Các thành phần này sẽ được hệ thống backend xử lý định dạng chuẩn theo loại văn bản.

Trả về **CHỈ** nội dung đã được làm sạch, loại bỏ tiêu đề và chữ ký, có đánh dấu tag chữ ký nếu có.

Nội dung cần xử lý:
{text_input}
"""

MAX_AI_INPUT_LENGTH = 15000
AI_RETRY_DELAY = 5

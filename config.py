# config.py
import os
from dotenv import load_dotenv
from docx.shared import Pt, Cm

load_dotenv()

# --- Biến cấu hình AI ---
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_API_KEY_HERE_IF_NO_ENV")
GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"
MAX_AI_INPUT_LENGTH = 15000
AI_RETRY_DELAY = 5

# --- PROMPT MỚI CHO AI (Ví dụ cho Nghị định - Cần tinh chỉnh và tạo thêm cho loại khác) ---
# Yêu cầu AI tạo phần body từ sau dòng "CHÍNH PHỦ" (lặp lại) đến trước "TM. CHÍNH PHỦ"
# AI phải tự áp dụng định dạng NĐ30 vào text trả về.
AI_PROMPT_TEMPLATE_NGHI_DINH = """
Bạn là một chuyên gia soạn thảo văn bản hành chính Việt Nam, tuân thủ tuyệt đối Nghị định 30/2020/NĐ-CP.
Nhiệm vụ của bạn là tạo ra phần **THÂN VĂN BẢN** cho một **Nghị định của Chính phủ** dựa trên yêu cầu và dữ liệu người dùng cung cấp ({user_input_data}).

**YÊU CẦU:**
1.  **Chỉ tạo phần THÂN VĂN BẢN:** Bắt đầu từ **sau** dòng "CHÍNH PHỦ" (dòng tên cơ quan lặp lại dưới tên Nghị định) và kết thúc **ngay trước** phần chữ ký ("TM. CHÍNH PHỦ").
2.  **KHÔNG BAO GỒM:** Không được tự ý thêm Quốc hiệu, Tiêu ngữ, Tên cơ quan ban hành ("CHÍNH PHỦ") ở đầu, Số/Ký hiệu, Địa danh/Ngày tháng, Tên loại ("NGHỊ ĐỊNH"), Trích yếu, Chữ ký, Nơi nhận.
3.  **ĐỊNH DẠNG TUYỆT ĐỐI THEO NĐ30:** Tự động áp dụng định dạng sau vào nội dung text bạn tạo ra:
    * **Font:** Toàn bộ nội dung sử dụng font 'Times New Roman'.
    * **Căn cứ:**
        * Căn lề: Đều hai bên (Justify).
        * Font: Cỡ chữ 14pt, *in nghiêng*.
        * Thụt lề dòng đầu: 1cm.
        * Giãn dòng: 1.5 lines.
        * Khoảng cách sau đoạn: 0pt.
        * Kết thúc: Dấu ';' cuối mỗi căn cứ (trừ dòng cuối cùng '.'). Thêm 1 dòng trống (paragraph trống) sau căn cứ cuối.
    * **Chương:**
        * Căn lề: Giữa (Center).
        * Font: Cỡ chữ 13pt, **in đậm**, IN HOA.
        * Cấu trúc: Dòng 1 "Chương [Số La Mã]", Dòng 2 "TÊN CHƯƠNG" (trong cùng 1 paragraph, dùng ngắt dòng).
        * Khoảng cách: Trước 12pt, Sau 6pt.
    * **Điều:**
        * Căn lề: Trái (Left).
        * Font: Cỡ chữ 13pt, **in đậm** (cả "Điều x." và tiêu đề điều).
        * Thụt lề dòng đầu: 0cm.
        * Khoảng cách: Trước 6pt, Sau 3pt.
        * Cấu trúc: "Điều [Số]. [Tiêu đề điều]"
    * **Khoản:**
        * Căn lề: Đều hai bên (Justify).
        * Font: Cỡ chữ 14pt.
        * Thụt lề toàn đoạn (left_indent): 1cm.
        * Thụt lề dòng đầu (first_line_indent): 0cm.
        * Giãn dòng: 1.5 lines.
        * Khoảng cách: Trước 3pt, Sau 3pt.
        * Cấu trúc: Bắt đầu bằng "[Số]." (ví dụ "1.", "2.").
    * **Điểm:**
        * Căn lề: Đều hai bên (Justify).
        * Font: Cỡ chữ 14pt.
        * Thụt lề toàn đoạn (left_indent): 1.5cm.
        * Thụt lề dòng đầu (first_line_indent): 0cm.
        * Giãn dòng: 1.5 lines.
        * Khoảng cách: Trước 3pt, Sau 3pt.
        * Cấu trúc: Bắt đầu bằng "[Chữ cái thường)]" (ví dụ "a)", "b)").
    * **Đoạn văn bản thường:**
        * Căn lề: Đều hai bên (Justify).
        * Font: Cỡ chữ 14pt.
        * Thụt lề dòng đầu: 1cm.
        * Giãn dòng: 1.5 lines.
        * Khoảng cách: Sau 6pt.
4.  **Nội dung:** Dựa vào {user_input_data} để viết nội dung chính xác, đúng văn phong hành chính.
5.  **Kết quả:** Chỉ trả về phần text của thân văn bản đã được định dạng theo mô tả. Không thêm giải thích.

**Dữ liệu người dùng cung cấp:**
{user_input_data}

Tạo phần thân văn bản cho Nghị định.
"""

# --- Hằng số định dạng NĐ30 Cơ bản ---
FONT_NAME = 'Times New Roman'

# Cỡ chữ cơ bản (Đã sửa lỗi import)
FONT_SIZE_11 = Pt(11)
FONT_SIZE_12 = Pt(12)
FONT_SIZE_13 = Pt(13)
FONT_SIZE_14 = Pt(14)

# Ánh xạ cỡ chữ theo thành phần để dễ quản lý và đảm bảo tồn tại
FONT_SIZE_DEFAULT = FONT_SIZE_14
FONT_SIZE_HEADER = FONT_SIZE_13
FONT_SIZE_TIEUNGU_DIADANH = FONT_SIZE_14
FONT_SIZE_SOKYHIEU = FONT_SIZE_13
FONT_SIZE_TITLE = FONT_SIZE_14      # Tên loại VB (NĐ, QĐ...) - Sửa lỗi import
FONT_SIZE_TRICHYEU = FONT_SIZE_14   # Trích yếu
FONT_SIZE_VV = FONT_SIZE_12         # V/v Công văn
FONT_SIZE_BODY = FONT_SIZE_14       # Nội dung chính
FONT_SIZE_CHUONG = FONT_SIZE_13     # Tiêu đề Chương
FONT_SIZE_DIEU = FONT_SIZE_13       # Tiêu đề Điều
FONT_SIZE_SIGNATURE_AUTH = FONT_SIZE_14 # Quyền hạn, Chức vụ ký
FONT_SIZE_SIGNATURE_NAME = FONT_SIZE_14 # Tên người ký
FONT_SIZE_RECIPIENT_LABEL = FONT_SIZE_12 # Chữ "Nơi nhận:"
FONT_SIZE_RECIPIENT_LIST = FONT_SIZE_11 # Danh sách nơi nhận
FONT_SIZE_SMALL = FONT_SIZE_11      # Cỡ chữ nhỏ khác nếu cần

# Lề trang chuẩn (Cm)
MARGIN_TOP = Cm(2.0)
MARGIN_BOTTOM = Cm(2.0)
MARGIN_LEFT_DEFAULT = Cm(3.0)
MARGIN_RIGHT_DEFAULT = Cm(1.5)

# Lề trang đặc biệt (Ví dụ)
MARGIN_LEFT_CONTRACT = Cm(3.0)
MARGIN_RIGHT_CONTRACT = Cm(1.5)

# Thụt lề dòng đầu tiên chuẩn
FIRST_LINE_INDENT = Cm(1.0)

# Giãn dòng chuẩn
LINE_SPACING_BODY = 1.5
LINE_SPACING_DEFAULT = 1.5 # Giữ lại nếu utils dùng
# config.py
import os
from dotenv import load_dotenv
from docx.shared import Pt, Cm

load_dotenv()

# --- Biến cấu hình AI và Ứng dụng ---
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_API_KEY_HERE_IF_NO_ENV")
GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"
AI_PROMPT_TEMPLATE = """
Bạn là một trợ lý biên tập viên tiếng Việt chuyên nghiệp cao cấp, cực kỳ cẩn thận và tuân thủ nghiêm ngặt yêu cầu.
Hãy đọc kỹ nội dung dưới đây và thực hiện chính xác các việc sau:

1.  **Sửa lỗi chính tả, ngữ pháp:** Đảm bảo câu chữ tiếng Việt chuẩn mực.
2.  **Tối ưu văn phong:** Loại bỏ từ ngữ thừa, Phần tiêu đề của văn bản (vì phần này đã được định dạng trong các văn bản chuẩn, không cần phần này nữa) câu lặp, diễn đạt khó hiểu. Đảm bảo văn phong mạch lạc, rõ ràng, trang trọng, phù hợp với ngữ cảnh văn bản hành chính/công việc. Giữ nguyên các thuật ngữ chuyên ngành nếu có.
3.  **Bảo toàn nội dung cốt lõi:** Giữ nguyên ý nghĩa gốc và các thông tin quan trọng như tên riêng (người, tổ chức, địa danh), số liệu, ngày tháng cụ thể có trong nội dung.
4.  **Xử lý tiêu đề không chuẩn:** Nếu phát hiện các phần text ở đầu văn bản rõ ràng không phải nội dung chính mà là tiêu đề tự thêm như “DỰ ÁN: ABC”, “CÔNG VĂN GỬI XYZ”, “VĂN BẢN THÔNG BÁO”, hoặc các dòng tương tự không thuộc cấu trúc chuẩn của thân văn bản — hãy **xóa bỏ hoàn toàn** các dòng tiêu đề không chuẩn đó.
5.  **XỬ LÝ TRIỆT ĐỂ PHẦN CHỮ KÝ/KẾT THÚC (Quan trọng nhất):**
    * Rà soát kỹ phần cuối của nội     * Nếu phát hiện bất kỳ dấu hiệu nào của phần chữ ký hoặc lời kết thúdung văn bản (Nếu trong văn bản quá dài như hợp đồng có nhiều chỗ cần ký phải xác định xóa cả những phần này nữa).
c thư, bao gồm nhưng không giới hạn ở các cụm từ/cấu trúc như:
        * Lời kết: "Trân trọng,", "Kính thư,", "Xin cảm ơn,", "Trân trọng kính chào,", "Kính đề nghị xem xét,", "Xin chân thành cảm ơn."
        * Chữ ký viết tắt/thay mặt: "TM.", "KT.", "TL.", "TUQ.", "Q." (Quyền), "PHỤ TRÁCH", "THỪA LỆNH", "THỪA ỦY QUYỀN"
        * Hướng dẫn ký: "(Ký tên, đóng dấu)", "(Ký, ghi rõ họ tên)", "(Đã ký)"
        * Placeholder: "[Chức vụ]", "[Tên người ký]", "(Họ tên, chức vụ)"
        * Chức danh (thường viết hoa): GIÁM ĐỐC, TỔNG GIÁM ĐỐC, CHỦ TỊCH, PHÓ CHỦ TỊCH, TRƯỞNG PHÒNG, PHÓ TRƯỞNG PHÒNG, HIỆU TRƯỞNG, PHÓ HIỆU TRƯỞNG, CHÁNH VĂN PHÒNG, THỦ TRƯỞNG ĐƠN VỊ, BAN GIÁM HIỆU,...
        * Tên riêng và/hoặc chức danh đứng một mình hoặc theo cụm ở cuối văn bản (ví dụ: Nguyễn Văn A, Phó Giám đốc Nguyễn Văn B).
        * Các dòng trống liên tiếp theo sau bởi tên/chức danh.
    * --> Hãy **XÓA SẠCH TOÀN BỘ CÁC DÒNG/CỤM TỪ ĐÓ KHỎI NỘI DUNG**. Mục tiêu là loại bỏ hoàn toàn mọi dấu vết của chữ ký hoặc lời kết thúc do người dùng tự thêm vào. Đảm bảo nội dung trả về kết thúc ngay sau phần nội dung chính cuối cùng, không còn sót lại bất kỳ yếu tố nào của chữ ký hay lời kết.
6.  **KHÔNG THÊM BẤT CỨ THỨ GÌ:** Tuyệt đối không tự ý thêm vào các thành phần như Quốc hiệu, Tiêu ngữ, Số ký hiệu, Ngày tháng ban hành, Nơi nhận, hoặc bất kỳ khối chữ ký nào. Việc này sẽ do hệ thống backend xử lý sau.
7.  **Kết quả trả về:** Trả về **CHỈ DUY NHẤT** phần nội dung văn bản đã được làm sạch và đã **XÓA HOÀN TOÀN** phần chữ ký/lời kết nếu có. Đảm bảo không có ký tự lạ, định dạng markdown (như ```) hoặc lời dẫn giải nào trong kết quả trả về.
Nội dung cần xử lý:
{text_input}
"""
MAX_AI_INPUT_LENGTH = 15000
AI_RETRY_DELAY = 5

# --- Hằng số định dạng chuẩn NĐ30 ---
FONT_NAME = 'Times New Roman'

# *** THÊM LẠI FONT_SIZE_DEFAULT ***
FONT_SIZE_DEFAULT = Pt(14) # Chọn 14pt làm mặc định vì NĐ30 cho phép 13-14

# Cỡ chữ chi tiết theo Phụ lục I, Mục V NĐ30
FONT_SIZE_HEADER_12 = Pt(12)
FONT_SIZE_HEADER_13 = Pt(13)
FONT_SIZE_MEDIUM_13 = Pt(13)
FONT_SIZE_MEDIUM_14 = Pt(14)
FONT_SIZE_TITLE_14 = Pt(14)
FONT_SIZE_VV_12 = Pt(12)
FONT_SIZE_VV_13 = Pt(13)
FONT_SIZE_BODY_13 = Pt(13)
FONT_SIZE_BODY_14 = Pt(14)
FONT_SIZE_SIGN_AUTH_13 = Pt(13)
FONT_SIZE_SIGN_AUTH_14 = Pt(14)
FONT_SIZE_SIGN_NAME_13 = Pt(13)
FONT_SIZE_SIGN_NAME_14 = Pt(14)
FONT_SIZE_RECIPIENT_LABEL_12 = Pt(12)
FONT_SIZE_RECIPIENT_LIST_11 = Pt(11)
FONT_SIZE_OTHER_11 = Pt(11)

# Lề trang chuẩn (Cm)
MARGIN_TOP = Cm(2.0)
MARGIN_BOTTOM = Cm(2.0)
MARGIN_LEFT_DEFAULT = Cm(3.0)
MARGIN_RIGHT_DEFAULT = Cm(1.5)

# Lề trang đặc biệt
MARGIN_LEFT_CONTRACT = Cm(3.0)
MARGIN_RIGHT_CONTRACT = Cm(1.5)

# Thụt lề dòng đầu tiên chuẩn
FIRST_LINE_INDENT = Cm(1.0)

# Giãn dòng chuẩn
LINE_SPACING_BODY = 1.5
LINE_SPACING_DEFAULT = 1.5 # Thêm hằng số này nếu utils dùng
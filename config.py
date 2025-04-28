# config.py
# Chỉ chứa các hằng số định dạng cần thiết theo NĐ30
from docx.shared import Pt, Cm

# Font chuẩn
FONT_NAME = 'Times New Roman'

# Cỡ chữ chuẩn theo Phụ lục I, Mục V NĐ30 (Ví dụ, cần kiểm tra lại)
# Header (QH, Tên CQ)
FONT_SIZE_HEADER_12 = Pt(12)
FONT_SIZE_HEADER_13 = Pt(13)
# Tiêu ngữ, Địa danh+TG, Số/KH
FONT_SIZE_MEDIUM_13 = Pt(13)
FONT_SIZE_MEDIUM_14 = Pt(14)
# Tên loại VB (QĐ, NQ,...) & Trích yếu
FONT_SIZE_TITLE_14 = Pt(14) # Thường 14pt
# V/v Công văn
FONT_SIZE_VV_12 = Pt(12)
FONT_SIZE_VV_13 = Pt(13)
# Nội dung chính (Điều, Khoản, Điểm, đoạn văn)
FONT_SIZE_BODY_13 = Pt(13)
FONT_SIZE_BODY_14 = Pt(14) # Phổ biến là 14pt cho dễ đọc
# Chữ ký (Quyền hạn, Chức vụ)
FONT_SIZE_SIGN_AUTH_13 = Pt(13)
FONT_SIZE_SIGN_AUTH_14 = Pt(14)
# Tên người ký
FONT_SIZE_SIGN_NAME_13 = Pt(13)
FONT_SIZE_SIGN_NAME_14 = Pt(14)
# Nơi nhận
FONT_SIZE_RECIPIENT_LABEL_12 = Pt(12)
FONT_SIZE_RECIPIENT_LIST_11 = Pt(11)
# Thành phần khác (Số trang, chỉ dẫn...)
FONT_SIZE_OTHER_11 = Pt(11)

# Lề trang chuẩn (Cm)
MARGIN_TOP = Cm(2.0)
MARGIN_BOTTOM = Cm(2.0)
MARGIN_LEFT = Cm(3.0)
MARGIN_RIGHT = Cm(1.5)

# Lề trang đặc biệt (Ví dụ Hợp đồng, có thể giống chuẩn)
MARGIN_LEFT_CONTRACT = Cm(3.0)
MARGIN_RIGHT_CONTRACT = Cm(1.5)

# Thụt lề dòng đầu tiên chuẩn
FIRST_LINE_INDENT = Cm(1.0) # Hoặc Cm(1.27) tùy theo quy định nội bộ

# Giãn dòng chuẩn (NĐ30: tối thiểu single, tối đa 1.5)
# Chọn 1.5 lines làm mặc định cho dễ đọc
LINE_SPACING_BODY = 1.5
# doc_formatter.py
import re
from docx import Document
from docx.shared import Cm
from config import MARGIN_TOP, MARGIN_BOTTOM, MARGIN_LEFT_DEFAULT, MARGIN_RIGHT_DEFAULT, MARGIN_LEFT_CONTRACT
# Import các hàm định dạng cụ thể từ thư mục formatters
from formatters import cong_van, quyet_dinh, chi_thi, thong_bao, ke_hoach
# (Import thêm các module formatter khác khi tạo chúng)

def identify_doc_type(title, body):
    """Nhận diện loại văn bản."""
    # (Giữ nguyên logic hàm identify_doc_type từ phiên bản trước)
    body_upper = body.upper()
    title_upper = title.upper()
    body_start_upper = body[:500].upper()

    if "QUYẾT ĐỊNH" in title_upper: return "QuyetDinh"
    if "CHỈ THỊ" in title_upper: return "ChiThi"
    if "THÔNG BÁO" in title_upper: return "ThongBao"
    if "KẾ HOẠCH" in title_upper: return "KeHoach"
    if "NGHỊ ĐỊNH" in title_upper: return "NghiDinh"
    if "QUY CHẾ" in title_upper: return "QuyChe"
    if "TỜ TRÌNH" in title_upper: return "ToTrinh"
    if "PHIẾU TRÌNH" in title_upper: return "PhieuTrinh"
    if "BÁO CÁO" in title_upper: return "BaoCao"
    if "GIẤY MỜI" in title_upper: return "GiayMoi"
    if "PHÁT BIỂU" in title_upper: return "PhatBieu"
    if "TIỂU LUẬN" in title_upper: return "TieuLuan"
    if "HỢP ĐỒNG" in title_upper: return "HopDong"

    if "QUYẾT ĐỊNH" in body_start_upper and "ĐIỀU" in body_upper: return "QuyetDinh"
    if "CHỈ THỊ" in body_start_upper: return "ChiThi"
    if "THÔNG BÁO" in body_start_upper: return "ThongBao"
    if "KẾ HOẠCH" in body_start_upper and ("MỤC ĐÍCH" in body_upper or "TỔ CHỨC THỰC HIỆN" in body_upper): return "KeHoach"
    if "NGHỊ ĐỊNH" in body_start_upper and "CHƯƠNG" in body_upper and "ĐIỀU" in body_upper: return "NghiDinh"

    return "CongVan" # Mặc định

# Dictionary ánh xạ loại văn bản với module định dạng
# *** Quan trọng: Cập nhật khi thêm module formatter mới ***
DOC_TYPE_FORMATTERS = {
    "CongVan": cong_van,
    "QuyetDinh": quyet_dinh,
    "ChiThi": chi_thi,
    "ThongBao": thong_bao,
    "KeHoach": ke_hoach,
    # "NghiDinh": nghi_dinh, # Ví dụ
    # "HopDong": hop_dong,
}

def apply_docx_formatting(data, doc_type):
    """Hàm chính điều phối việc tạo và định dạng Document."""
    document = Document()

    # 1. Thiết lập lề trang cơ bản
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM
    # Lề trái có thể thay đổi tùy loại (vd: Hợp đồng)
    section.left_margin = MARGIN_LEFT_CONTRACT if doc_type == "HopDong" else MARGIN_LEFT_DEFAULT
    section.right_margin = MARGIN_RIGHT_DEFAULT

    # 2. Lấy hàm định dạng phù hợp
    # Nếu không tìm thấy loại cụ thể, dùng Công văn làm mặc định
    formatter_module = DOC_TYPE_FORMATTERS.get(doc_type, cong_van)
    print(f"Sử dụng module định dạng: formatters.{formatter_module.__name__}")

    # 3. Gọi hàm format chính của module đó
    # Mỗi module formatter cần có hàm chuẩn là format(document, data)
    if hasattr(formatter_module, 'format'):
        formatter_module.format(document, data)
    else:
        print(f"Lỗi: Module {formatter_module.__name__} thiếu hàm format(document, data)")
        # Có thể fallback về định dạng công văn mặc định ở đây
        cong_van.format(document, data)

    print("Định dạng Word hoàn tất.")
    return document
# doc_formatter.py
import re
from docx import Document
from docx.shared import Cm
from config import MARGIN_TOP, MARGIN_BOTTOM, MARGIN_LEFT_DEFAULT, MARGIN_RIGHT_DEFAULT, MARGIN_LEFT_CONTRACT

# --- Định nghĩa PlaceholderFormatter ở ngoài ---
class PlaceholderFormatter:
    def format(self, document, data):
        print(f"Warning: PlaceholderFormatter activated. Using basic paragraph.")
        document.add_paragraph(data.get('body', ''))

# --- Khối Import: Import TẤT CẢ các formatter hiện có ---
try:
    from formatters import ban_ghi_nho, ban_thoa_thuan, bang_tot_nghiep, bao_cao, bien_ban, \
                           chi_thi, chuong_trinh, cong_dien, cong_van, de_an, \
                           de_cuong_mh, don_nhap_hoc, du_an, giao_trinh, giay_bao_trung_tuyen, \
                           giay_gioi_thieu, giay_moi, giay_nghi_phep, giay_uy_quyen, giay_xac_nhan_sv, \
                           hop_dong, huong_dan, huong_dan_hs, ke_hoach, luat, \
                           nghi_dinh, nghi_dinh_qppl, nghi_quyet, nghi_quyet_qh, phap_lenh, \
                           phat_bieu, phieu, phieu_trinh, phuong_an, quy_che, \
                           quy_che_ts, quy_dinh, quy_dinh_nt, quyet_dinh, quyet_dinh_ts, \
                           quyet_dinh_ttg, thong_bao, thong_bao_nt, thong_bao_ts, thong_cao, \
                           thong_tu, thu_cong, tieu_luan, to_trinh

    print("Tất cả formatters đã được import thành công.")

except ImportError as e:
    print(f"LỖI IMPORT FORMATTERS: {e}. Sẽ sử dụng Placeholder cho các module lỗi.")
    # Gán fallback cho TẤT CẢ nếu import lỗi
    ban_ghi_nho = ban_thoa_thuan = bang_tot_nghiep = bao_cao = bien_ban = \
    chi_thi = chuong_trinh = cong_dien = cong_van = de_an = \
    de_cuong_mh = don_nhap_hoc = du_an = giao_trinh = giay_bao_trung_tuyen = \
    giay_gioi_thieu = giay_moi = giay_nghi_phep = giay_uy_quyen = giay_xac_nhan_sv = \
    hop_dong = huong_dan = huong_dan_hs = ke_hoach = luat = \
    nghi_dinh = nghi_dinh_qppl = nghi_quyet = nghi_quyet_qh = phap_lenh = \
    phat_bieu = phieu = phieu_trinh = phuong_an = quy_che = \
    quy_che_ts = quy_dinh = quy_dinh_nt = quyet_dinh = quyet_dinh_ts = \
    quyet_dinh_ttg = thong_bao = thong_bao_nt = thong_bao_ts = thong_cao = \
    thong_tu = thu_cong = tieu_luan = to_trinh = PlaceholderFormatter()


# --- HÀM NHẬN DIỆN: Bao gồm tất cả loại văn bản hiện có ---
def identify_doc_type(title, body):
    body_upper = body.upper()
    title_upper = title.upper()
    body_start_upper = body[:800].upper()

    if "BẰNG TỐT NGHIỆP" in title_upper or "CHỨNG CHỈ" in title_upper: return "BangTotNghiep"
    if "GIÁO TRÌNH" in title_upper: return "GiaoTrinh"
    if "ĐƠN XIN NHẬP HỌC" in title_upper or "PHIẾU ĐĂNG KÝ NHẬP HỌC" in title_upper: return "DonNhapHoc"
    # if "THỜI KHÓA BIỂU" in title_upper: return "ThoiKhoaBieu" # Đã xóa
    # if "BẢNG ĐIỂM" in title_upper or "PHIẾU ĐIỂM" in title_upper: return "BangDiem" # Đã xóa
    if "ĐỀ CƯƠNG MÔN HỌC" in title_upper or "ĐỀ CƯƠNG CHI TIẾT" in title_upper: return "DeCuongMH"
    if "QUY ĐỊNH" in title_upper and ("HỌC SINH" in body_upper or "SINH VIÊN" in body_upper or "NỘI QUY" in body_upper): return "QuyDinhNT"
    if "THÔNG BÁO" in title_upper and ("NHÀ TRƯỜNG" in body_start_upper or "KHOA" in body_start_upper or "SINH VIÊN" in body_upper): return "ThongBaoNT"
    if "THÔNG BÁO TUYỂN SINH" in title_upper or ("TUYỂN SINH" in title_upper and "THÔNG BÁO" in title_upper): return "ThongBaoTS"
    if "QUY CHẾ TUYỂN SINH" in title_upper: return "QuyCheTS"
    if "HƯỚNG DẪN" in title_upper and ("HỒ SƠ" in title_upper or "NHẬP HỌC" in title_upper or "TUYỂN SINH" in title_upper): return "HuongDanHS"
    if "GIẤY BÁO TRÚNG TUYỂN" in title_upper or "GIẤY BÁO NHẬP HỌC" in title_upper: return "GiayBaoTrungTuyen"
    if "QUYẾT ĐỊNH" in title_upper and ("TRÚNG TUYỂN" in title_upper or "CÔNG NHẬN SINH VIÊN" in title_upper): return "QuyetDinhTS"
    if "GIẤY XÁC NHẬN" in title_upper and ("SINH VIÊN" in title_upper or "NGƯỜI HỌC" in title_upper): return "GiayXacNhanSV"

    if title_upper.startswith("LUẬT") or title_upper.startswith("BỘ LUẬT"): return "Luat"
    if "NGHỊ QUYẾT" in title_upper and ("QUỐC HỘI" in body_start_upper or "QUỐC HỘI KHÓA" in body_start_upper): return "NghiQuyetQH"
    if title_upper.startswith("PHÁP LỆNH"): return "PhapLenh"
    if title_upper.startswith("NGHỊ ĐỊNH") and ("CHÍNH PHỦ" in body_start_upper or "CĂN CỨ LUẬT" in body_start_upper) and ("CHƯƠNG" in body_upper or "ĐIỀU" in body_upper): return "NghiDinhQPPL"
    if title_upper.startswith("QUYẾT ĐỊNH") and ("THỦ TƯỚNG CHÍNH PHỦ" in body_start_upper or "CĂN CỨ LUẬT" in body_start_upper) and ("ĐIỀU" in body_upper): return "QuyetDinhTTg"
    if title_upper.startswith("THÔNG TƯ"): return "ThongTu"

    if "GIẤY ỦY QUYỀN" in title_upper: return "GiayUyQuyen"
    if "GIẤY GIỚI THIỆU" in title_upper: return "GiayGioiThieu"
    if "GIẤY NGHỈ PHÉP" in title_upper or "ĐƠN XIN NGHỈ PHÉP" in title_upper: return "GiayNghiPhep"
    if title_upper.startswith("PHIẾU GỬI") or title_upper.startswith("PHIẾU CHUYỂN") or title_upper.startswith("PHIẾU BÁO"): return "Phieu"
    if "THƯ CÔNG" in title_upper or "CÔNG THƯ" in title_upper: return "ThuCong"
    if "HỢP ĐỒNG" in title_upper or ("BÊN A" in body_start_upper and "BÊN B" in body_start_upper and "ĐIỀU KHOẢN" in body_upper): return "HopDong"
    if "THÔNG CÁO" in title_upper: return "ThongCao"
    if "PHƯƠNG ÁN" in title_upper: return "PhuongAn"
    if "DỰ ÁN" in title_upper: return "DuAn"
    if "CÔNG ĐIỆN" in title_upper or "CÔNG ĐIỆN" in body_start_upper: return "CongDien"
    if "BẢN GHI NHỚ" in title_upper or "MEMORANDUM OF UNDERSTANDING" in title_upper: return "BanGhiNho"
    if "BẢN THỎA THUẬN" in title_upper or "AGREEMENT" in title_upper: return "BanThoaThuan"
    if "NGHỊ QUYẾT" in title_upper and "QUYẾT NGHỊ:" in body_upper: return "NghiQuyet"
    if "QUY ĐỊNH" in title_upper: return "QuyDinh"
    if "HƯỚNG DẪN" in title_upper: return "HuongDan"
    if "CHƯƠNG TRÌNH" in title_upper: return "ChuongTrinh"
    if "BIÊN BẢN" in title_upper: return "BienBan"
    if "ĐỀ ÁN" in title_upper: return "DeAn"
    if "QUYẾT ĐỊNH" in title_upper: return "QuyetDinh"
    if "CHỈ THỊ" in title_upper: return "ChiThi"
    if "THÔNG BÁO" in title_upper: return "ThongBao"
    if "KẾ HOẠCH" in title_upper: return "KeHoach"
    if "NGHỊ ĐỊNH" in title_upper: return "NghiDinh" # NĐ Hành chính
    if "QUY CHẾ" in title_upper: return "QuyChe"
    if "TỜ TRÌNH" in title_upper: return "ToTrinh"
    if "PHIẾU TRÌNH" in title_upper: return "PhieuTrinh"
    if "BÁO CÁO" in title_upper: return "BaoCao"
    if "GIẤY MỜI" in title_upper: return "GiayMoi"
    if "PHÁT BIỂU" in title_upper: return "PhatBieu"
    if "TIỂU LUẬN" in title_upper: return "TieuLuan"

    # Nhận diện dựa trên nội dung
    if "GIÁO TRÌNH" in body_start_upper and "CHƯƠNG" in body_upper: return "GiaoTrinh"
    if "ỦY QUYỀN CHO" in body_upper and "NỘI DUNG ỦY QUYỀN" in body_upper: return "GiayUyQuyen"
    if "TRÂN TRỌNG GIỚI THIỆU ÔNG/BÀ" in body_upper and "ĐƯỢC CỬ ĐẾN" in body_upper: return "GiayGioiThieu"
    if "KÍNH GỬI" in body_start_upper and "XIN NGHỈ PHÉP" in body_upper and "LÝ DO" in body_upper: return "GiayNghiPhep"
    if "KÍNH GỬI" in body_start_upper and "TRÂN TRỌNG" in body_upper and body.count('\n') < 20: return "ThuCong"
    if "THÔNG CÁO BÁO CHÍ" in body_start_upper: return "ThongCao"
    if "PHƯƠNG ÁN" in body_start_upper and ("MỤC TIÊU" in body_upper or "GIẢI PHÁP" in body_upper): return "PhuongAn"
    if "DỰ ÁN" in body_start_upper and ("CHỦ ĐẦU TƯ" in body_upper or "TỔNG MỨC ĐẦU TƯ" in body_upper): return "DuAn"
    if "BÊN A" in body_start_upper and "BÊN B" in body_start_upper and "THỐNG NHẤT GHI NHỚ" in body_upper: return "BanGhiNho"
    if "BÊN A" in body_start_upper and "BÊN B" in body_start_upper and "CÙNG THỎA THUẬN" in body_upper: return "BanThoaThuan"
    if "QUYẾT NGHỊ:" in body_upper: return "NghiQuyet"
    if "QUY ĐỊNH" in body_start_upper and "ĐIỀU" in body_upper: return "QuyDinh"
    if "HƯỚNG DẪN THỰC HIỆN" in body_start_upper: return "HuongDan"
    if "CHƯƠNG TRÌNH CÔNG TÁC" in body_start_upper or ("MỤC TIÊU" in body_upper and "NỘI DUNG" in body_upper and "TỔ CHỨC THỰC HIỆN" in body_upper): return "ChuongTrinh"
    if "BIÊN BẢN" in body_start_upper and "THỜI GIAN" in body_upper and "ĐỊA ĐIỂM" in body_upper and "THÀNH PHẦN" in body_upper : return "BienBan"
    if "ĐỀ ÁN" in body_start_upper and "SỰ CẦN THIẾT" in body_upper and "MỤC TIÊU" in body_upper and "GIẢI PHÁP" in body_upper: return "DeAn"
    if "QUYẾT ĐỊNH" in body_start_upper and "ĐIỀU" in body_upper and "THỦ TƯỚNG CHÍNH PHỦ" not in body_start_upper and "TRÚNG TUYỂN" not in body_upper : return "QuyetDinh"
    if "CHỈ THỊ" in body_start_upper: return "ChiThi"
    if "THÔNG BÁO" in body_start_upper and "TUYỂN SINH" not in body_upper: return "ThongBao"
    if "KẾ HOẠCH" in body_start_upper and ("MỤC ĐÍCH" in body_upper or "TỔ CHỨC THỰC HIỆN" in body_upper): return "KeHoach"
    if "NGHỊ ĐỊNH" in body_start_upper and ("CHÍNH PHỦ" not in body_start_upper or "CHƯƠNG" not in body_upper): return "NghiDinh" # NĐ Hành chính

    return None


# --- DICTIONARY ÁNH XẠ: Bao gồm tất cả các formatter hiện có ---
DOC_TYPE_FORMATTERS = {
    # QPPL & Cơ bản
    "Luat": luat, "NghiQuyetQH": nghi_quyet_qh, "PhapLenh": phap_lenh,
    "NghiDinhQPPL": nghi_dinh_qppl, "QuyetDinhTTg": quyet_dinh_ttg, "ThongTu": thong_tu,
    # Hành chính thông thường & Mẫu đơn
    "CongVan": cong_van, "QuyetDinh": quyet_dinh, "ChiThi": chi_thi, "ThongBao": thong_bao,
    "KeHoach": ke_hoach, "NghiQuyet": nghi_quyet, "QuyDinh": quy_dinh, "HuongDan": huong_dan,
    "ChuongTrinh": chuong_trinh, "BienBan": bien_ban, "DeAn": de_an, "ThongCao": thong_cao,
    "PhuongAn": phuong_an, "DuAn": du_an, "CongDien": cong_dien, "BanGhiNho": ban_ghi_nho,
    "BanThoaThuan": ban_thoa_thuan, "GiayUyQuyen": giay_uy_quyen, "GiayGioiThieu": giay_gioi_thieu,
    "GiayNghiPhep": giay_nghi_phep, "Phieu": phieu, "ThuCong": thu_cong, "HopDong": hop_dong,
    "NghiDinh": nghi_dinh, "QuyChe": quy_che, "ToTrinh": to_trinh, "PhieuTrinh": phieu_trinh,
    "BaoCao": bao_cao, "GiayMoi": giay_moi, "PhatBieu": phat_bieu, "TieuLuan": tieu_luan,
    # Giáo dục & Tuyển sinh
    "ThongBaoTS": thong_bao_ts, "QuyCheTS": quy_che_ts, "HuongDanHS": huong_dan_hs,
    "GiayBaoTrungTuyen": giay_bao_trung_tuyen, "QuyetDinhTS": quyet_dinh_ts,
    "GiayXacNhanSV": giay_xac_nhan_sv, "DonNhapHoc": don_nhap_hoc,
    # "ThoiKhoaBieu": thoi_khoa_bieu, # Đã xóa
    # "BangDiem": bang_diem, # Đã xóa
    "DeCuongMH": de_cuong_mh, "QuyDinhNT": quy_dinh_nt,
    "ThongBaoNT": thong_bao_nt, "BangTotNghiep": bang_tot_nghiep, "GiaoTrinh": giao_trinh,
}


# --- HÀM ĐIỀU PHỐI: Đã sửa lỗi NameError và bao gồm các loại lề đặc biệt ---
def apply_docx_formatting(data, recognized_doc_type, intended_doc_type):
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM

    formatter_module = None
    doc_type_for_filename = None
    is_placeholder_used = False # Cờ để biết có dùng placeholder không

    # Ưu tiên 1: Thử dùng loại văn bản nhận diện được
    if recognized_doc_type:
        formatter_module = DOC_TYPE_FORMATTERS.get(recognized_doc_type)
        if formatter_module and not isinstance(formatter_module, PlaceholderFormatter):
            doc_type_for_filename = recognized_doc_type
            print(f"Sử dụng formatter cho loại NHẬN DIỆN: formatters.{getattr(formatter_module, '__name__', 'N/A')}")
        else:
            if formatter_module is None:
                 print(f"Formatter cho loại NHẬN DIỆN '{recognized_doc_type}' không tìm thấy trong DOC_TYPE_FORMATTERS.")
            else: # Là PlaceholderFormatter
                 print(f"Formatter cho loại NHẬN DIỆN '{recognized_doc_type}' là Placeholder (có thể do import lỗi trước đó).")
                 is_placeholder_used = True
            formatter_module = None # Reset để thử loại dự định

    # Ưu tiên 2: Nếu không thành công với loại nhận diện, thử dùng loại dự định
    if formatter_module is None and intended_doc_type:
        formatter_module = DOC_TYPE_FORMATTERS.get(intended_doc_type)
        if formatter_module and not isinstance(formatter_module, PlaceholderFormatter):
            doc_type_for_filename = intended_doc_type
            print(f"Sử dụng formatter cho loại DỰ ĐỊNH: formatters.{getattr(formatter_module, '__name__', 'N/A')}")
            is_placeholder_used = False # Đã tìm thấy formatter thật
        else:
            if formatter_module is None:
                 print(f"Formatter cho loại DỰ ĐỊNH '{intended_doc_type}' không tìm thấy trong DOC_TYPE_FORMATTERS.")
            else: # Là PlaceholderFormatter
                 print(f"Formatter cho loại DỰ ĐỊNH '{intended_doc_type}' là Placeholder (có thể do import lỗi trước đó).")
                 is_placeholder_used = True
            formatter_module = None # Reset để dùng fallback

    # Ưu tiên 3: Nếu cả hai đều không được, dùng fallback
    if formatter_module is None:
        if is_placeholder_used:
            print("Sử dụng formatter Placeholder do lỗi import trước đó.")
            formatter_module = PlaceholderFormatter()
            doc_type_for_filename = intended_doc_type if intended_doc_type in DOC_TYPE_FORMATTERS else "Fallback"
        else:
            print("Sử dụng formatter mặc định: formatters.cong_van")
            formatter_module = cong_van # Đảm bảo cong_van đã được import thành công
            doc_type_for_filename = intended_doc_type if intended_doc_type in DOC_TYPE_FORMATTERS and not isinstance(DOC_TYPE_FORMATTERS.get(intended_doc_type), PlaceholderFormatter) else "CongVan"


    # Lấy tên module cuối cùng được chọn
    final_formatter_name = getattr(formatter_module, '__name__', 'PlaceholderFormatter')
    doc_type_used_for_formatting = doc_type_for_filename if doc_type_for_filename else ("CongVan" if not is_placeholder_used else "Fallback")


    # Thiết lập lề trang và hướng trang dựa trên loại formatter cuối cùng
    if doc_type_used_for_formatting in ["BanGhiNho", "BanThoaThuan", "HopDong"]:
        section.left_margin = MARGIN_LEFT_CONTRACT
        section.right_margin = MARGIN_RIGHT_DEFAULT
    # elif doc_type_used_for_formatting in ["ThoiKhoaBieu", "BangDiem"]: # Đã xóa
    #      section.left_margin = Cm(1.5)
    #      section.right_margin = Cm(1.5)
    elif doc_type_used_for_formatting == "BangTotNghiep":
         section.orientation = 1 # WD_ORIENTATION.LANDSCAPE = 1
         section.page_width = Cm(29.7)
         section.page_height = Cm(21.0)
         section.left_margin = Cm(1.5)
         section.right_margin = Cm(1.5)
         section.top_margin = Cm(1.5)
         section.bottom_margin = Cm(1.5)
    elif doc_type_used_for_formatting == "TieuLuan":
        # Lề tiểu luận thường khác biệt
        section.left_margin = Cm(3.0) # Lề trái rộng hơn
        section.right_margin = Cm(2.0) # Lề phải hẹp hơn
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
    else: # Lề mặc định cho các loại văn bản khác
        section.left_margin = MARGIN_LEFT_DEFAULT
        section.right_margin = MARGIN_RIGHT_DEFAULT


    # Gọi hàm format của module đã chọn
    if hasattr(formatter_module, 'format'):
        try:
            formatter_module.format(document, data)
        except Exception as e:
             print(f"Lỗi khi chạy formatter cho {doc_type_used_for_formatting} ({final_formatter_name}): {e}")
             import traceback
             traceback.print_exc()
             # Thêm thông báo lỗi vào tài liệu để dễ debug
             document.add_paragraph(f"--- Lỗi định dạng văn bản ---")
             document.add_paragraph(f"Loại văn bản: {doc_type_used_for_formatting}")
             document.add_paragraph(f"Lỗi: {e}")
             document.add_paragraph(f"Xem chi tiết trong log của server.")
    else:
        print(f"Lỗi: Module {final_formatter_name} không có hàm format(document, data).")
        document.add_paragraph(f"Lỗi: Không thể định dạng văn bản loại '{doc_type_used_for_formatting}'.")


    print("Định dạng Word hoàn tất.")
    # Trả về document và loại dùng cho tên file
    return document, doc_type_used_for_formatting # Trả về loại đã thực sự dùng để định dạng
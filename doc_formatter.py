# doc_formatter.py
import re
from docx import Document
from docx.shared import Cm
from config import MARGIN_TOP, MARGIN_BOTTOM, MARGIN_LEFT_DEFAULT, MARGIN_RIGHT_DEFAULT, MARGIN_LEFT_CONTRACT

# --- PHẦN IMPORT: Đảm bảo import TẤT CẢ các module formatters bạn đã tạo ---
try:
    from formatters import cong_van, quyet_dinh, chi_thi, thong_bao, ke_hoach, \
                           nghi_quyet, quy_dinh, huong_dan, chuong_trinh, bien_ban, de_an, \
                           thong_cao, phuong_an, du_an, cong_dien, ban_ghi_nho, ban_thoa_thuan, \
                           giay_uy_quyen, giay_gioi_thieu, giay_nghi_phep, phieu, thu_cong, hop_dong, \
                           luat, nghi_quyet_qh, phap_lenh, nghi_dinh_qppl, quyet_dinh_ttg, thong_tu, \
                           thong_bao_ts, quy_che_ts, huong_dan_hs, giay_bao_trung_tuyen, quyet_dinh_ts, giay_xac_nhan_sv, \
                           don_nhap_hoc, thoi_khoa_bieu, bang_diem, de_cuong_mh, quy_dinh_nt, thong_bao_nt, \
                           bang_tot_nghiep, giao_trinh, \
                           quy_che, to_trinh, phieu_trinh, bao_cao, giay_moi, phat_bieu, tieu_luan, \
                           nghi_dinh # Đảm bảo tên file và import khớp nhau

except ImportError as e:
    print(f"Lỗi import formatters: {e}. Đảm bảo các file formatter tồn tại trong thư mục formatters/ và tên file đúng.")
    class PlaceholderFormatter:
        def format(self, document, data):
            print(f"Warning: Formatter not found. Using basic paragraph.")
            document.add_paragraph(data.get('body', ''))

    # Gán fallback cho TẤT CẢ nếu import lỗi
    cong_van = quyet_dinh = chi_thi = thong_bao = ke_hoach = \
    nghi_quyet = quy_dinh = huong_dan = chuong_trinh = bien_ban = de_an = \
    thong_cao = phuong_an = du_an = cong_dien = ban_ghi_nho = ban_thoa_thuan = \
    giay_uy_quyen = giay_gioi_thieu = giay_nghi_phep = phieu = thu_cong = hop_dong = \
    luat = nghi_quyet_qh = phap_lenh = nghi_dinh_qppl = quyet_dinh_ttg = thong_tu = \
    thong_bao_ts = quy_che_ts = huong_dan_hs = giay_bao_trung_tuyen = quyet_dinh_ts = giay_xac_nhan_sv = \
    don_nhap_hoc = thoi_khoa_bieu = bang_diem = de_cuong_mh = quy_dinh_nt = thong_bao_nt = \
    bang_tot_nghiep = giao_trinh = \
    quy_che = to_trinh = phieu_trinh = bao_cao = giay_moi = phat_bieu = tieu_luan = \
    nghi_dinh = PlaceholderFormatter()

# --- HÀM NHẬN DIỆN: Giữ nguyên logic nhận diện ---
def identify_doc_type(title, body):
    body_upper = body.upper()
    title_upper = title.upper()
    body_start_upper = body[:800].upper()

    if "BẰNG TỐT NGHIỆP" in title_upper or "CHỨNG CHỈ" in title_upper: return "BangTotNghiep"
    if "GIÁO TRÌNH" in title_upper: return "GiaoTrinh"
    if "ĐƠN XIN NHẬP HỌC" in title_upper or "PHIẾU ĐĂNG KÝ NHẬP HỌC" in title_upper: return "DonNhapHoc"
    if "THỜI KHÓA BIỂU" in title_upper: return "ThoiKhoaBieu"
    if "BẢNG ĐIỂM" in title_upper or "PHIẾU ĐIỂM" in title_upper: return "BangDiem"
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
    if "NGHỊ ĐỊNH" in title_upper: return "NghiDinh"
    if "QUY CHẾ" in title_upper: return "QuyChe"
    if "TỜ TRÌNH" in title_upper: return "ToTrinh"
    if "PHIẾU TRÌNH" in title_upper: return "PhieuTrinh"
    if "BÁO CÁO" in title_upper: return "BaoCao"
    if "GIẤY MỜI" in title_upper: return "GiayMoi"
    if "PHÁT BIỂU" in title_upper: return "PhatBieu"
    if "TIỂU LUẬN" in title_upper: return "TieuLuan"

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
    if "NGHỊ ĐỊNH" in body_start_upper and ("CHÍNH PHỦ" not in body_start_upper or "CHƯƠNG" not in body_upper): return "NghiDinh"

    return None


# --- DICTIONARY ÁNH XẠ: Đảm bảo tất cả các key và value đều đúng ---
DOC_TYPE_FORMATTERS = {
    "Luat": luat, "NghiQuyetQH": nghi_quyet_qh, "PhapLenh": phap_lenh,
    "NghiDinhQPPL": nghi_dinh_qppl, "QuyetDinhTTg": quyet_dinh_ttg, "ThongTu": thong_tu,
    "CongVan": cong_van, "QuyetDinh": quyet_dinh, "ChiThi": chi_thi, "ThongBao": thong_bao,
    "KeHoach": ke_hoach, "NghiQuyet": nghi_quyet, "QuyDinh": quy_dinh, "HuongDan": huong_dan,
    "ChuongTrinh": chuong_trinh, "BienBan": bien_ban, "DeAn": de_an, "ThongCao": thong_cao,
    "PhuongAn": phuong_an, "DuAn": du_an, "CongDien": cong_dien, "BanGhiNho": ban_ghi_nho,
    "BanThoaThuan": ban_thoa_thuan, "GiayUyQuyen": giay_uy_quyen, "GiayGioiThieu": giay_gioi_thieu,
    "GiayNghiPhep": giay_nghi_phep, "Phieu": phieu, "ThuCong": thu_cong, "HopDong": hop_dong,
    "NghiDinh": nghi_dinh, "QuyChe": quy_che, "ToTrinh": to_trinh, "PhieuTrinh": phieu_trinh,
    "BaoCao": bao_cao, "GiayMoi": giay_moi, "PhatBieu": phat_bieu, "TieuLuan": tieu_luan,
    "ThongBaoTS": thong_bao_ts, "QuyCheTS": quy_che_ts, "HuongDanHS": huong_dan_hs,
    "GiayBaoTrungTuyen": giay_bao_trung_tuyen, "QuyetDinhTS": quyet_dinh_ts,
    "GiayXacNhanSV": giay_xac_nhan_sv, "DonNhapHoc": don_nhap_hoc, "ThoiKhoaBieu": thoi_khoa_bieu,
    "BangDiem": bang_diem, "DeCuongMH": de_cuong_mh, "QuyDinhNT": quy_dinh_nt,
    "ThongBaoNT": thong_bao_nt, "BangTotNghiep": bang_tot_nghiep, "GiaoTrinh": giao_trinh,
}

# --- HÀM ĐIỀU PHỐI: Giữ nguyên logic xử lý fallback ---
def apply_docx_formatting(data, recognized_doc_type, intended_doc_type):
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_BOTTOM

    formatter_module = None
    doc_type_for_filename = None

    if recognized_doc_type:
        formatter_module = DOC_TYPE_FORMATTERS.get(recognized_doc_type)
        if formatter_module and not isinstance(formatter_module, PlaceholderFormatter):
            doc_type_for_filename = recognized_doc_type
            print(f"Sử dụng formatter cho loại NHẬN DIỆN: formatters.{getattr(formatter_module, '__name__', 'N/A')}")
        else:
            print(f"Formatter cho loại NHẬN DIỆN '{recognized_doc_type}' không hợp lệ hoặc không tìm thấy.")
            formatter_module = None

    if formatter_module is None and intended_doc_type:
        formatter_module = DOC_TYPE_FORMATTERS.get(intended_doc_type)
        if formatter_module and not isinstance(formatter_module, PlaceholderFormatter):
            doc_type_for_filename = intended_doc_type
            print(f"Sử dụng formatter cho loại DỰ ĐỊNH: formatters.{getattr(formatter_module, '__name__', 'N/A')}")
        else:
            print(f"Formatter cho loại DỰ ĐỊNH '{intended_doc_type}' cũng không hợp lệ hoặc không tìm thấy.")
            formatter_module = None

    if formatter_module is None:
        print("Sử dụng formatter mặc định: formatters.cong_van")
        formatter_module = cong_van
        doc_type_for_filename = intended_doc_type if intended_doc_type in DOC_TYPE_FORMATTERS and not isinstance(DOC_TYPE_FORMATTERS.get(intended_doc_type), PlaceholderFormatter) else "CongVan"

    final_formatter_name = getattr(formatter_module, '__name__', 'cong_van')
    doc_type_used_for_formatting = doc_type_for_filename if doc_type_for_filename else "CongVan"

    if doc_type_used_for_formatting in ["BanGhiNho", "BanThoaThuan", "HopDong"]:
        section.left_margin = MARGIN_LEFT_CONTRACT
        section.right_margin = MARGIN_RIGHT_DEFAULT
    elif doc_type_used_for_formatting in ["ThoiKhoaBieu", "BangDiem"]:
         section.left_margin = Cm(1.5)
         section.right_margin = Cm(1.5)
    elif doc_type_used_for_formatting == "BangTotNghiep":
         section.orientation = 1
         section.page_width = Cm(29.7)
         section.page_height = Cm(21.0)
         section.left_margin = Cm(1.5)
         section.right_margin = Cm(1.5)
         section.top_margin = Cm(1.5)
         section.bottom_margin = Cm(1.5)
    elif doc_type_used_for_formatting == "TieuLuan":
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
    else:
        section.left_margin = MARGIN_LEFT_DEFAULT
        section.right_margin = MARGIN_RIGHT_DEFAULT

    if hasattr(formatter_module, 'format'):
        try:
            formatter_module.format(document, data)
        except Exception as e:
             print(f"Lỗi khi chạy formatter cho {doc_type_used_for_formatting} ({final_formatter_name}): {e}")
             import traceback
             traceback.print_exc()
             document.add_paragraph(f"Lỗi định dạng văn bản loại '{doc_type_used_for_formatting}': {e}")
    else:
        print(f"Lỗi: Module {final_formatter_name} thiếu hàm format(document, data).")
        document.add_paragraph(f"Lỗi: Không thể định dạng văn bản loại '{doc_type_used_for_formatting}'.")


    print("Định dạng Word hoàn tất.")
    return document, doc_type_for_filename
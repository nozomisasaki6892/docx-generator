# formatters/nghi_dinh.py
import re
import time
import traceback # Thêm để báo lỗi chi tiết
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from utils import (
    set_paragraph_format,
    set_run_format,
    add_run_with_format,
    add_paragraph_with_text
)
from common_elements import (
    add_quoc_hieu_tieu_ngu,
    add_ten_co_quan_ban_hanh,
    add_so_ky_hieu,
    add_dia_danh_thoi_gian,
    add_signature_block,
    add_recipient_list
)
# Import hằng số chuẩn hóa
from config import (
    FONT_NAME,
    FONT_SIZE_HEADER_13, FONT_SIZE_MEDIUM_14, FONT_SIZE_MEDIUM_13,
    FONT_SIZE_TITLE_14, FONT_SIZE_BODY_14, FONT_SIZE_BODY_13,
    FONT_SIZE_SIGN_AUTH_14, FONT_SIZE_SIGN_NAME_14,
    FONT_SIZE_RECIPIENT_LABEL_12, FONT_SIZE_RECIPIENT_LIST_11,
    FIRST_LINE_INDENT, LINE_SPACING_BODY
)

# Cỡ chữ mặc định cho nội dung Nghị định
DEFAULT_BODY_FONT_SIZE = FONT_SIZE_BODY_14

def format(document, data):
    """Định dạng văn bản theo thể thức Nghị định của Chính phủ (QPPL)."""
    print("--- Bắt đầu định dạng Nghị định ---")
    try:
        # --- 1. Header ---
        print("Đang tạo header...")
        header_table = document.add_table(rows=1, cols=2)
        header_table.autofit = False
        header_table.allow_autofit = False
        header_table.columns[0].width = Inches(2.9)
        header_table.columns[1].width = Inches(3.3)
        cell_left = header_table.cell(0, 0)
        cell_right = header_table.cell(0, 1)
        cell_left._element.clear_content()
        cell_right._element.clear_content()

        add_ten_co_quan_ban_hanh(cell_left, None, "CHÍNH PHỦ")
        so_van_ban = data.get("decree_number_only", "...")
        ky_hieu_van_ban = data.get("decree_symbol", "NĐ-CP")
        add_so_ky_hieu(cell_left, so_van_ban, ky_hieu_van_ban)

        add_quoc_hieu_tieu_ngu(cell_right)
        dia_danh = data.get("issuing_location", "Hà Nội")
        try:
            ngay = int(data.get("issuing_day", time.strftime("%d")))
            thang = int(data.get("issuing_month", time.strftime("%m")))
            nam = int(data.get("issuing_year", time.strftime("%Y")))
        except ValueError:
             ngay, thang, nam = int(time.strftime("%d")), int(time.strftime("%m")), int(time.strftime("%Y"))
             print("Cảnh báo: Không thể parse ngày/tháng/năm từ data, dùng ngày hiện tại.")
        add_dia_danh_thoi_gian(cell_right, dia_danh, ngay, thang, nam)
        print("Tạo header xong.")
        document.add_paragraph()

        # --- 2. Tên loại và Trích yếu ---
        print("Đang thêm tên loại và trích yếu...")
        nghi_dinh_label = "NGHỊ ĐỊNH"
        add_paragraph_with_text(document, nghi_dinh_label, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(6), space_after=Pt(6), size=FONT_SIZE_TITLE_14, bold=True, uppercase=True)

        trich_yeu = data.get("title", "...") # Lấy title làm trích yếu
        # Chuẩn hóa title nếu nó bắt đầu bằng "Nghị định..."
        if trich_yeu.lower().startswith("nghị định"):
            trich_yeu = trich_yeu[len("nghị định"):].strip()
        if trich_yeu.lower().startswith("về việc"):
             trich_yeu = trich_yeu[len("về việc"):].strip()

        paragraph_trichyeu = add_paragraph_with_text(document, trich_yeu, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(6), size=FONT_SIZE_TITLE_14, bold=True)
        paragraph_line = document.add_paragraph()
        set_paragraph_format(paragraph_line, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
        add_run_with_format(paragraph_line, "________", size=FONT_SIZE_TITLE_14, bold=True)
        print("Thêm tên loại và trích yếu xong.")

        # --- 3. Cơ quan ban hành (lặp lại) ---
        print("Đang thêm tên cơ quan ban hành (lặp lại)...")
        add_paragraph_with_text(document, "CHÍNH PHỦ", alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12), size=FONT_SIZE_MEDIUM_13, bold=True, uppercase=True)
        print("Thêm tên cơ quan ban hành xong.")

        # --- 4. Căn cứ ---
        print("Đang xử lý căn cứ...")
        can_cu_list = data.get("can_cu", [])
        if can_cu_list:
            for item in can_cu_list:
                stripped_item = item.strip()
                if stripped_item:
                    add_paragraph_with_text(
                        document, stripped_item, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                        first_line_indent=FIRST_LINE_INDENT, line_spacing=LINE_SPACING_BODY,
                        space_after=Pt(0), size=DEFAULT_BODY_FONT_SIZE, italic=True
                    )
            document.add_paragraph()
        print("Xử lý căn cứ xong.")

        # --- 5. Nội dung Body ---
        print("Đang xử lý nội dung body...")
        body_content = data.get("body", "")
        body_lines = body_content.strip().split('\n')

        for line_index, line in enumerate(body_lines):
            stripped_line = line.strip()
            if not stripped_line:
                continue

            # Quan trọng: Bỏ qua nếu dòng này là lặp lại tên loại văn bản
            if stripped_line.upper() == nghi_dinh_label:
                print(f"  INFO: Bỏ qua dòng tiêu đề lặp lại '{stripped_line}' ở dòng {line_index+1}")
                continue

            paragraph = document.add_paragraph()
            align = WD_ALIGN_PARAGRAPH.JUSTIFY
            left_indent = Cm(0)
            first_indent = FIRST_LINE_INDENT
            is_bold = False
            size = DEFAULT_BODY_FONT_SIZE
            space_before = Pt(0)
            space_after = Pt(6)
            line_spacing = LINE_SPACING_BODY
            text_to_add = stripped_line
            is_heading = False # Cờ đánh dấu là tiêu đề mục

            # Logic nhận diện (Cần test kỹ với dữ liệu AI thực tế)
            match_chuong = re.match(r'^(Chương\s+[IVXLCDM]+)\s*(.*?)$', stripped_line, re.IGNORECASE)
            match_dieu = re.match(r'^(Điều\s+\d+)\.\s*(.*)', stripped_line, re.IGNORECASE)
            match_khoan = re.match(r'^(\d+)\.\s+(.*)', stripped_line)
            match_diem = re.match(r'^([a-z])\)\s+(.*)', stripped_line)

            if match_chuong:
                is_heading = True
                chuong_num = match_chuong.group(1)
                chuong_title = match_chuong.group(2).strip().upper()
                align = WD_ALIGN_PARAGRAPH.CENTER
                first_indent = Cm(0)
                is_bold = True
                size = FONT_SIZE_BODY_13 # Cỡ chữ nhỏ hơn chút cho Chương
                space_before = Pt(12)
                space_after = Pt(6)
                text_to_add = f"{chuong_num}\n{chuong_title}"
                paragraph.clear() # Xóa text gốc để add lại với format
                add_run_with_format(paragraph, text_to_add, size=size, bold=is_bold)

            elif match_dieu:
                is_heading = True
                dieu_num_title = match_dieu.group(1)
                dieu_content = match_dieu.group(2).strip()
                align = WD_ALIGN_PARAGRAPH.LEFT # Điều căn trái
                first_indent = Cm(0)
                is_bold = True
                size = FONT_SIZE_BODY_13
                space_before = Pt(6)
                space_after = Pt(3)
                paragraph.clear()
                add_run_with_format(paragraph, f"{dieu_num_title}. ", size=size, bold=is_bold)
                add_run_with_format(paragraph, dieu_content, size=size, bold=is_bold) # Tiêu đề đậm
                text_to_add = None

            elif match_khoan:
                khoan_num = match_khoan.group(1)
                khoan_content = match_khoan.group(2).strip()
                left_indent = FIRST_LINE_INDENT
                first_indent = Cm(0)
                is_bold = False
                size = DEFAULT_BODY_FONT_SIZE
                space_before = Pt(3)
                space_after = Pt(3)
                paragraph.clear()
                add_run_with_format(paragraph, f"{khoan_num}. {khoan_content}", size=size, bold=is_bold)
                text_to_add = None

            elif match_diem:
                diem_marker = match_diem.group(1)
                diem_content = match_diem.group(2).strip()
                left_indent = FIRST_LINE_INDENT + Cm(0.5) # Thụt thêm
                first_indent = Cm(0)
                is_bold = False
                size = DEFAULT_BODY_FONT_SIZE
                space_before = Pt(3)
                space_after = Pt(3)
                paragraph.clear()
                add_run_with_format(paragraph, f"{diem_marker}) {diem_content}", size=size, bold=is_bold)
                text_to_add = None

            # Chỉ áp dụng định dạng đoạn cho các dòng đã được xử lý hoặc dòng thường
            if text_to_add is not None and not is_heading: # Thêm text nếu là đoạn thường
                 add_run_with_format(paragraph, text_to_add, size=size, bold=is_bold)

            set_paragraph_format(paragraph, alignment=align, left_indent=left_indent,
                                 first_line_indent=first_indent, line_spacing=line_spacing,
                                 space_before=space_before, space_after=space_after)

        print("Xử lý nội dung body xong.")

        # --- 6. Chữ ký ---
        print("Đang thêm chữ ký...")
        signer_authority = data.get("authority_signer", "TM. CHÍNH PHỦ")
        signer_title = data.get("signer_title", "THỦ TƯỚNG")
        signer_name = data.get("signer_name", "[Tên Thủ tướng]")
        add_signature_block(document, authority_signer=signer_authority, signer_title=signer_title, signer_name=signer_name)
        print("Thêm chữ ký xong.")

        # --- 7. Nơi nhận ---
        print("Đang thêm nơi nhận...")
        default_recipients = [
            "- Ban Bí thư Trung ương Đảng;", "- Thủ tướng, các Phó Thủ tướng Chính phủ;",
            "- Các bộ, cơ quan ngang bộ, cơ quan thuộc Chính phủ;",
            "- HĐND, UBND các tỉnh, thành phố trực thuộc trung ương;",
            "- Văn phòng Trung ương và các Ban của Đảng;", "- Văn phòng Tổng Bí thư;",
            "- Văn phòng Chủ tịch nước;", "- Hội đồng Dân tộc và các Ủy ban của Quốc hội;",
            "- Văn phòng Quốc hội;", "- Tòa án nhân dân tối cao;",
            "- Viện kiểm sát nhân dân tối cao;", "- Kiểm toán Nhà nước;",
            "- VPCP: BTCN, các PCN, Trợ lý TTg, TGĐ Cổng TTĐT, các Vụ, Cục, đơn vị trực thuộc, Công báo;",
            "- Lưu: VT, [Ký hiệu đơn vị soạn thảo]."
        ]
        recipients = data.get("recipients", default_recipients)
        add_recipient_list(document, recipients)
        print("Thêm nơi nhận xong.")

    except Exception as error:
        print(f"!!!!!!!! LỖI TRONG QUÁ TRÌNH ĐỊNH DẠNG NGHỊ ĐỊNH !!!!!!!!!!")
        print(traceback.format_exc())
        # Thêm thông báo lỗi vào tài liệu để dễ nhận biết
        try:
            document.add_paragraph(f"--- LỖI ĐỊNH DẠNG: {error} ---")
            document.add_paragraph(traceback.format_exc())
        except Exception as inner_error:
             print(f"Lỗi khi ghi thông báo lỗi vào doc: {inner_error}")

    print("--- Định dạng Nghị định hoàn tất ---")
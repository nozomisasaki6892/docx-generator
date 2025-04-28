# common_elements.py
import time
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import set_paragraph_format, set_run_format, add_run_with_format, add_paragraph_with_text

# Lưu ý: Cần import các hằng số cỡ chữ từ config.py đã được chuẩn hóa theo NĐ30
# Ví dụ: FONT_SIZE_12, FONT_SIZE_13, FONT_SIZE_14, FONT_SIZE_11,...
# Tạm thời dùng Pt trực tiếp hoặc giả định hằng số
FONT_SIZE_11 = Pt(11)
FONT_SIZE_12 = Pt(12)
FONT_SIZE_13 = Pt(13)
FONT_SIZE_14 = Pt(14)

def add_quoc_hieu_tieu_ngu(table_cell):
    """Thêm Quốc hiệu và Tiêu ngữ vào ô bảng được chỉ định (Ô số 1)."""
    # Quốc hiệu
    paragraph_qh = table_cell.add_paragraph()
    set_paragraph_format(paragraph_qh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(paragraph_qh, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", size=FONT_SIZE_13, bold=True) # Cỡ 12-13

    # Tiêu ngữ
    paragraph_tn = table_cell.add_paragraph()
    set_paragraph_format(paragraph_tn, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    add_run_with_format(paragraph_tn, "Độc lập - Tự do - Hạnh phúc", size=FONT_SIZE_14, bold=True) # Cỡ 13-14

    # Đường kẻ dưới Tiêu ngữ
    paragraph_line = table_cell.add_paragraph()
    set_paragraph_format(paragraph_line, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0)) # Giảm space_after
    # Điều chỉnh độ dài dòng kẻ phù hợp
    add_run_with_format(paragraph_line, "_______________", size=FONT_SIZE_14, bold=True)


def add_ten_co_quan_ban_hanh(table_cell, ten_co_quan_chu_quan, ten_co_quan_ban_hanh):
    """Thêm Tên cơ quan chủ quản (nếu có) và Tên cơ quan ban hành (Ô số 2)."""
    # Cơ quan chủ quản (nếu có)
    if ten_co_quan_chu_quan:
        paragraph_cqcq = table_cell.add_paragraph()
        set_paragraph_format(paragraph_cqcq, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
        # Cỡ chữ 12-13, đứng, không đậm
        add_run_with_format(paragraph_cqcq, ten_co_quan_chu_quan.upper(), size=FONT_SIZE_13, bold=False)

    # Tên cơ quan ban hành
    paragraph_cqbh = table_cell.add_paragraph()
    set_paragraph_format(paragraph_cqbh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    # Cỡ chữ 12-13, đứng, đậm
    add_run_with_format(paragraph_cqbh, ten_co_quan_ban_hanh.upper(), size=FONT_SIZE_13, bold=True)

    # Đường kẻ dưới tên CQBH
    paragraph_line = table_cell.add_paragraph()
    set_paragraph_format(paragraph_line, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0))
    # Độ dài 1/3-1/2 dòng chữ CQBH, đậm
    add_run_with_format(paragraph_line, "________", size=FONT_SIZE_13, bold=True) # Điều chỉnh độ dài


def add_so_ky_hieu(table_cell, so_van_ban, ky_hieu_van_ban):
    """Thêm Số, ký hiệu văn bản vào ô bảng (Ô số 3)."""
    paragraph_skh = table_cell.add_paragraph()
    set_paragraph_format(paragraph_skh, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0)) # Giảm space_after
    # Chữ "Số:" đứng, thường, cỡ 13
    add_run_with_format(paragraph_skh, f"Số: {so_van_ban}/", size=FONT_SIZE_13, bold=False)
    # Ký hiệu đứng, hoa, cỡ 13
    add_run_with_format(paragraph_skh, ky_hieu_van_ban, size=FONT_SIZE_13, bold=False, uppercase=True)


def add_dia_danh_thoi_gian(table_cell, dia_danh, ngay, thang, nam):
    """Thêm Địa danh, thời gian ban hành vào ô bảng (Ô số 4)."""
    paragraph_ddtg = table_cell.add_paragraph()
    set_paragraph_format(paragraph_ddtg, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(0)) # Giảm space_after
    # Cỡ 13-14, nghiêng
    thoi_gian_str = f"ngày {ngay:02d} tháng {thang:02d} năm {nam}"
    add_run_with_format(paragraph_ddtg, f"{dia_danh}, {thoi_gian_str}", size=FONT_SIZE_13, italic=True)


def add_signature_block(
    document, # Hoặc table_cell nếu đặt trong bảng
    authority_signer=None, # "TM.", "KT.", "TL.", "TUQ.", "Q."
    signer_title="",
    signer_name="",
    signer_note=None, # "(Đã ký)"
    alignment=WD_ALIGN_PARAGRAPH.RIGHT, # Mặc định căn phải
    use_table=True # Dùng table để dễ căn chỉnh hơn
):
    """Thêm khối chữ ký (Ô số 7a, 7b, 7c)."""

    if use_table:
        # Tạo bảng 1 hàng 2 cột để đẩy khối ký sang phải
        sig_table = document.add_table(rows=1, cols=2)
        sig_table.autofit = False
        sig_table.allow_autofit = False
        # Điều chỉnh độ rộng cột tùy ý, ví dụ cột trái trống, cột phải chứa chữ ký
        sig_table.columns[0].width = Inches(3.0) # Cột trống
        sig_table.columns[1].width = Inches(3.3) # Cột chữ ký
        cell_sig = sig_table.cell(0, 1)
        cell_sig._element.clear_content() # Xóa mọi nội dung cũ trong ô
        paragraph_container = cell_sig # Thêm paragraph vào ô này
        alignment = WD_ALIGN_PARAGRAPH.CENTER # Căn giữa trong ô bên phải
    else:
        paragraph_container = document # Thêm paragraph trực tiếp vào document
        # Cần tính toán left_indent nếu không dùng table để đẩy sang phải

    # Thẩm quyền ký (nếu có) - Ô 7a
    if authority_signer:
        paragraph_auth = paragraph_container.add_paragraph()
        set_paragraph_format(paragraph_auth, alignment=alignment, space_after=Pt(0))
        # Cỡ 13-14, Đứng, Đậm, IN HOA
        add_run_with_format(paragraph_auth, authority_signer.upper(), size=FONT_SIZE_14, bold=True)

    # Chức vụ người ký - Ô 7a
    paragraph_title = paragraph_container.add_paragraph()
    set_paragraph_format(paragraph_title, alignment=alignment, space_after=Pt(0))
    # Cỡ 13-14, Đứng, Đậm, IN HOA
    add_run_with_format(paragraph_title, signer_title.upper(), size=FONT_SIZE_14, bold=True)

    # Ghi chú ký (nếu có)
    if signer_note:
        paragraph_note = paragraph_container.add_paragraph()
        set_paragraph_format(paragraph_note, alignment=alignment, space_after=Pt(0))
        add_run_with_format(paragraph_note, signer_note, size=FONT_SIZE_11, italic=True) # Ví dụ cỡ 11

    # Khoảng trống ký tên - Ô 7c (giả lập bằng dòng trống)
    paragraph_space = paragraph_container.add_paragraph("\n\n\n\n") # Khoảng 4-5 dòng trống
    set_paragraph_format(paragraph_space, alignment=alignment, space_after=Pt(0))

    # Tên người ký - Ô 7b
    paragraph_name = paragraph_container.add_paragraph()
    set_paragraph_format(paragraph_name, alignment=alignment, space_after=Pt(0))
    # Cỡ 13-14, Đứng, Đậm
    add_run_with_format(paragraph_name, signer_name, size=FONT_SIZE_14, bold=True)


def add_recipient_list(document, recipients):
    """Thêm danh sách nơi nhận (Ô số 9b)."""
    paragraph_label = document.add_paragraph()
    # Sát lề trái, cỡ 12, nghiêng, đậm
    set_paragraph_format(paragraph_label, alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=Pt(6), space_after=Pt(0))
    add_run_with_format(paragraph_label, "Nơi nhận:", size=FONT_SIZE_12, bold=True, italic=True)

    if not recipients:
        recipients = ["- Như Điều ...;", "- Lưu: VT, ..."] # Mặc định nếu list rỗng

    for recipient in recipients:
        paragraph_rec = document.add_paragraph()
        # Đầu dòng có gạch ngang, sát lề trái, cỡ 11, đứng
        set_paragraph_format(
            paragraph_rec,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            left_indent=Cm(0.7),  # Thụt vào để gạch đầu dòng thẳng hàng
            first_line_indent=Cm(-0.7), # Hanging indent
            space_before=Pt(0),
            space_after=Pt(0),
            line_spacing=1.0 # Giãn dòng đơn cho nơi nhận
        )
        add_run_with_format(paragraph_rec, recipient, size=FONT_SIZE_11, bold=False)
# doc_formatter.py (Kiến trúc mới)
import traceback
from docx import Document
from docx.shared import Pt, Cm, Inches # Đảm bảo import đủ
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Import utils và common_elements đã được refactor
import utils
import common_elements
# Import các hằng số cần thiết từ config đã refactor
from config import (
    FONT_SIZE_TITLE, FONT_SIZE_TRICHYEU, FONT_SIZE_HEADER, # Sử dụng tên hằng số mới
    DEFAULT_BODY_FONT_SIZE # Mặc định nếu cần
)

# --- Bỏ hoàn toàn logic identify_doc_type và DOC_TYPE_FORMATTERS ---

def create_formatted_document(data, intended_doc_type):
    """
    Tạo tài liệu Word hoàn chỉnh theo kiến trúc mới:
    Python tạo khung (header, signature, recipients) + AI tạo body đã định dạng.
    """
    print(f"--- doc_formatter START (New Arch): Creating document for type '{intended_doc_type}' ---", flush=True)
    document = Document() # Tạo document mới
    print("  DEBUG doc_formatter: Initial Document object created.", flush=True)

    try:
        # 1. Áp dụng lề trang chuẩn (hoặc lề cụ thể nếu cần)
        print(f"  DEBUG doc_formatter: Applying margins for type '{intended_doc_type}'...", flush=True)
        # Ví dụ: Có thể thêm logic kiểm tra intended_doc_type để áp dụng lề khác
        # if intended_doc_type == "HopDong":
        #     utils.apply_specific_margins(document, ...)
        # else:
        utils.apply_standard_margins(document) # Mặc định dùng lề chuẩn

        # 2. Thêm Header chuẩn (dùng hàm từ common_elements)
        # Đảm bảo 'data' chứa đủ các key cần thiết được chuẩn bị từ app.py
        print("  DEBUG doc_formatter: Calling common_elements.add_header_elements...", flush=True)
        common_elements.add_header_elements(document, data)
        print("  DEBUG doc_formatter: Finished common_elements.add_header_elements.", flush=True)

        # 3. Thêm Tên loại và Trích yếu
        print("  DEBUG doc_formatter: Adding Ten loai & Trich yeu...", flush=True)
        # Lấy tên loại từ intended_doc_type hoặc data, cần chuẩn hóa ở app.py
        doc_type_label_map = {
            "NghiDinh": "NGHỊ ĐỊNH",
            "QuyetDinh": "QUYẾT ĐỊNH",
            "CongVan": "CÔNG VĂN", # Thường không có tên loại riêng nhưng trích yếu quan trọng
            # Thêm các loại khác vào đây
            "ToTrinh": "TỜ TRÌNH",
            "BaoCao": "BÁO CÁO",
            "ThongBao": "THÔNG BÁO",
            # ...
        }
        doc_type_label = doc_type_label_map.get(intended_doc_type, intended_doc_type or "[LOẠI VĂN BẢN]").upper()

        trich_yeu = data.get("title", "[TRÍCH YẾU NỘI DUNG]")
        # Xử lý trích yếu nếu nó chứa tên loại (ví dụ: "Nghị định về việc ABC")
        if trich_yeu.lower().startswith(doc_type_label.lower()[:len(doc_type_label)//2]): # Kiểm tra nửa đầu tên loại
             try:
                  trich_yeu = trich_yeu.split(" ", 1)[1] # Bỏ từ đầu tiên
             except IndexError:
                  pass # Giữ nguyên nếu không tách được
        if trich_yeu.lower().startswith("về việc"):
             trich_yeu = trich_yeu[len("về việc"):].strip()

        # Chỉ thêm tên loại nếu không phải Công văn (Công văn dùng V/v)
        if intended_doc_type != "CongVan":
            utils.add_paragraph_with_text(
                document, doc_type_label, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                space_before=Pt(6), space_after=Pt(6), size=FONT_SIZE_TITLE, bold=True
            )
            paragraph_trichyeu = utils.add_paragraph_with_text(
                document, trich_yeu, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                space_after=Pt(6), size=FONT_SIZE_TRICHYEU, bold=True
            )
            # Dòng kẻ dưới trích yếu
            paragraph_line = document.add_paragraph()
            utils.set_paragraph_format(paragraph_line, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(12))
            utils.add_run_with_format(paragraph_line, "________", size=FONT_SIZE_TRICHYEU, bold=True)
        else:
             # Xử lý riêng cho Công văn (Thêm V/v)
             paragraph_vv = utils.add_paragraph_with_text(
                document, f"V/v: {trich_yeu}", alignment=WD_ALIGN_PARAGRAPH.CENTER,
                space_before=Pt(6), space_after=Pt(12), size=FONT_SIZE_VV, bold=False # V/v thường không đậm
             )
        print("  DEBUG doc_formatter: Ten loai & Trich yeu added.", flush=True)

        # 4. Thêm Tên cơ quan ban hành lặp lại (Ví dụ: cho Nghị định QPPL)
        # Cần chuẩn hóa logic này dựa trên intended_doc_type
        if intended_doc_type in ["NghiDinh", "NghiDinhQPPL"]: # Chỉ ví dụ
             print("  DEBUG doc_formatter: Adding repeated issuer name...", flush=True)
             issuer_name = data.get("issuing_org", "[TÊN CƠ QUAN]").upper()
             utils.add_paragraph_with_text(
                 document, issuer_name, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                 space_after=Pt(12), size=FONT_SIZE_HEADER, bold=True, uppercase=True
             )
             print("  DEBUG doc_formatter: Repeated issuer name added.", flush=True)

        # 5. Thêm phần thân văn bản từ AI
        print("  DEBUG doc_formatter: Adding AI-generated body content...", flush=True)
        ai_body_text = data.get('body', '[Lỗi: Không có nội dung body từ AI]')
        if ai_body_text:
            # Cách 1: Thêm từng dòng như paragraph mới (Đơn giản nhất)
            body_lines = ai_body_text.strip().split('\n')
            for line in body_lines:
                if line.strip(): # Bỏ qua dòng hoàn toàn trống
                    # Thêm paragraph với nội dung từ AI, không áp dụng định dạng Python
                    # Giả định AI đã trả về text đúng định dạng
                    document.add_paragraph(line)
            print(f"  DEBUG doc_formatter: Added {len(body_lines)} lines from AI body (simple).", flush=True)
            # ---
            # Cách 2: (Nâng cao hơn - nếu cần xử lý markdown/tag từ AI)
            # Ví dụ nếu AI trả về **Điều 1.** ABC
            # for line in body_lines:
            #     if line.strip():
            #         p = document.add_paragraph()
            #         if line.startswith("**Điều"):
            #              # Xóa markdown, áp dụng style Điều bằng utils
            #              cleaned_line = line.replace("**","")
            #              utils.add_run_with_format(p, cleaned_line, bold=True, size=FONT_SIZE_DIEU)
            #              utils.set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, ...)
            #         elif line.startswith("1."):
            #              # Xử lý Khoản...
            #              utils.add_run_with_format(p, line, size=FONT_SIZE_BODY)
            #              utils.set_paragraph_format(p, left_indent=FIRST_LINE_INDENT, ...)
            #         else: # Đoạn thường
            #              utils.add_run_with_format(p, line, size=FONT_SIZE_BODY)
            #              utils.set_paragraph_format(p, first_line_indent=FIRST_LINE_INDENT, ...)
            # print(f"  DEBUG doc_formatter: Added {len(body_lines)} lines from AI body (parsing).", flush=True)
            # --- Chọn 1 trong 2 cách trên ---
        else:
            document.add_paragraph("[Lỗi: Nội dung body trống]")
            print("  WARNING doc_formatter: AI body content is empty!", flush=True)

        # 6. Thêm khối chữ ký chuẩn
        print("  DEBUG doc_formatter: Calling common_elements.add_signature_block...", flush=True)
        common_elements.add_signature_block(document, data)
        print("  DEBUG doc_formatter: Finished common_elements.add_signature_block.", flush=True)

        # 7. Thêm khối nơi nhận chuẩn
        print("  DEBUG doc_formatter: Calling common_elements.add_recipient_list...", flush=True)
        common_elements.add_recipient_list(document, data)
        print("  DEBUG doc_formatter: Finished common_elements.add_recipient_list.", flush=True)

        print(f"--- doc_formatter END (New Arch): Document created successfully for type '{intended_doc_type}' ---", flush=True)
        return document, intended_doc_type # Trả về intended_doc_type cho tên file

    except Exception as e:
        print(f"!!!!!!!! ERROR in create_formatted_document (New Arch) !!!!!!!!!!", flush=True)
        error_trace = traceback.format_exc()
        print(error_trace, flush=True)
        # Tạo document lỗi
        error_doc = Document()
        utils.apply_standard_margins(error_doc)
        error_doc.add_paragraph(f"--- LỖI NGHIÊM TRỌNG KHI TẠO VĂN BẢN ({intended_doc_type}) ---")
        error_doc.add_paragraph(f"Error: {str(e)}")
        error_doc.add_paragraph("Traceback:")
        for line in error_trace.splitlines():
            error_doc.add_paragraph(line)
        return error_doc, f"Error_{intended_doc_type}"
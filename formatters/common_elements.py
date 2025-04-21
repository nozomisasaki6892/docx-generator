# formatters/common_elements.py
import time
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from utils import set_paragraph_format, set_run_format, add_run_with_format
from config import FONT_SIZE_HEADER, FONT_SIZE_DEFAULT, FONT_SIZE_SMALL

# (Chuyển các hàm format_basic_header, format_signature_block, format_recipient_list từ app.py cũ vào đây)

def format_basic_header(document, data, doc_type):
    """Tạo phần header chuẩn cho văn bản hành chính."""
    issuing_org = data.get("issuing_org", "TÊN CƠ QUAN/TỔ CHỨC").upper()
    doc_number = data.get("doc_number", "Số:       /...")
    issuing_location = data.get("issuing_location", "Hà Nội")
    current_date_str = time.strftime(f"ngày %d tháng %m năm %Y")

    header_table = document.add_table(rows=1, cols=2)
    # (Giữ nguyên logic table header từ phiên bản trước)
    # ... (Copy code format_basic_header từ app.py cũ vào đây) ...

def format_signature_block(document, data):
    """Tạo khối chữ ký chuẩn."""
    signer_title = data.get("signer_title", "CHỨC VỤ").upper()
    signer_name = data.get("signer_name", "Người Ký")
    authority_signer = data.get("authority_signer", None)

    p_sig_block = document.add_paragraph()
    # (Giữ nguyên logic signature block từ phiên bản trước)
    # ... (Copy code format_signature_block từ app.py cũ vào đây) ...

def format_recipient_list(document, data):
    """Tạo phần nơi nhận."""
    recipients = data.get("recipients", ["- Như trên;", "- Lưu: VT, ..."])

    p_nhan_label = document.add_paragraph()
    # (Giữ nguyên logic recipient list từ phiên bản trước)
    # ... (Copy code format_recipient_list từ app.py cũ vào đây) ...
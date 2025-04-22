# app.py
import io
import os
import re
import uuid
import traceback
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from config import *
from ai_processor import call_gemini_api_for_cleanup
from doc_formatter import identify_doc_type as recognize_document_type
from doc_formatter import apply_docx_formatting as format_word_document
from doc_formatter import DOC_TYPE_FORMATTERS
from doc_formatter import PlaceholderFormatter  # 🛠 Thêm dòng này để tránh lỗi

app = Flask(__name__)
CORS(app)

# --- Hàm xử lý tag chữ ký ---
def extract_and_update_signature(body_cleaned, data):
    extracted_title = None
    extracted_name = None

    # Tìm và trích xuất chức danh
    pos_match = re.search(r'\[SIGNATURE_POSITION\](.*?)\[/SIGNATURE_POSITION\]', body_cleaned, re.DOTALL)
    if pos_match:
        extracted_title = pos_match.group(1).strip()
        if extracted_title: # Chỉ cập nhật nếu trích xuất được nội dung
            print(f"AI extracted signer_title: '{extracted_title}'")
            data['signer_title'] = extracted_title # Ghi đè giá trị từ data gốc

    # Tìm và trích xuất tên người ký
    name_match = re.search(r'\[SIGNATURE_NAME\](.*?)\[/SIGNATURE_NAME\]', body_cleaned, re.DOTALL)
    if name_match:
        extracted_name = name_match.group(1).strip()
        if extracted_name: # Chỉ cập nhật nếu trích xuất được nội dung
             print(f"AI extracted signer_name: '{extracted_name}'")
             data['signer_name'] = extracted_name # Ghi đè giá trị từ data gốc

    # Xóa các thẻ tag khỏi body_cleaned
    body_after_tag_removal = re.sub(r'\[SIGNATURE_POSITION\].*?\[/SIGNATURE_POSITION\]', '', body_cleaned, flags=re.DOTALL).strip()
    body_after_tag_removal = re.sub(r'\[SIGNATURE_NAME\].*?\[/SIGNATURE_NAME\]', '', body_after_tag_removal, flags=re.DOTALL).strip()

    return body_after_tag_removal, data
# --- Kết thúc hàm xử lý tag ---


@app.route("/generate", methods=["POST"])
def generate_docx_route():
    try:
        data = request.get_json(force=True)
        title = data.get("title", "Văn bản")
        body_original = data.get("body", "")
        intended_doc_type = data.get("intended_doc_type", None)

        if not body_original:
            return jsonify({"error": "Nội dung văn bản không được để trống"}), 400

        print(f"\nNhận yêu cầu tạo: '{title}'")
        if intended_doc_type:
            print(f"Loại văn bản người dùng dự định: {intended_doc_type}")
        else:
            print("Không có thông tin loại văn bản dự định từ frontend.")

        recognized_doc_type = recognize_document_type(title, body_original)
        if recognized_doc_type:
             print(f"Loại văn bản nhận diện được: {recognized_doc_type}")
        else:
             print("Không nhận diện được loại văn bản cụ thể từ nội dung.")

        use_ai_cleanup = data.get("use_ai", True)
        body_cleaned = body_original # Khởi tạo

        if use_ai_cleanup:
            print("Đang gọi AI để làm sạch nội dung...")
            body_cleaned = call_gemini_api_for_cleanup(body_original)
            print("Nội dung đã qua AI làm sạch.")

            # --- Xử lý thẻ Tag chữ ký SAU KHI AI làm sạch ---
            body_cleaned, data = extract_and_update_signature(body_cleaned, data)
            # --- Kết thúc xử lý thẻ Tag ---

        else:
            # body_cleaned = body_original # Đã khởi tạo ở trên
            print("Bỏ qua bước AI làm sạch.")

        # Gán nội dung cuối cùng (đã xử lý tag nếu có) vào data['body']
        data['body'] = body_cleaned

        # Gọi hàm định dạng với data đã được cập nhật (nếu có)
        document, doc_type_for_filename = format_word_document(data, recognized_doc_type, intended_doc_type)

        output_stream = io.BytesIO()
        document.save(output_stream)
        output_stream.seek(0)
        safe_title = re.sub(r'[\\/*?:"<>|]', "", title).replace(" ", "_")
        # Sử dụng doc_type_for_filename (loại thực sự dùng để format) cho tên file
        filename = f"{safe_title}_{doc_type_for_filename}_{uuid.uuid4().hex[:6]}.docx"
        print(f"Tạo file: {filename}")

        return send_file(
            output_stream,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print(f"Lỗi nghiêm trọng trong route /generate: {e}")
        traceback.print_exc()
        return jsonify({"error": f"Lỗi server: {str(e)}"}), 500

@app.route("/", methods=["GET"])
def home():
    # Lấy danh sách các loại được hỗ trợ từ keys của DOC_TYPE_FORMATTERS
    # Kiểm tra xem value có phải là PlaceholderFormatter không nếu cần độ chính xác cao hơn
    supported_types = [k for k, v in DOC_TYPE_FORMATTERS.items() if not isinstance(v, PlaceholderFormatter)]
    return jsonify({
        "message": "API tạo văn bản Word chuẩn VN (v5.0 - Final) sẵn sàng tại /generate", # Cập nhật version message
        "supported_document_types": sorted(supported_types)
     })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    # Chạy với debug=False cho production
    app.run(host="0.0.0.0", port=port, debug=False)
# app.py
import io
import os
import re
import uuid
import traceback
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
# Import config trước để đảm bảo biến môi trường được load nếu cần
from config import * # Có thể import cụ thể hơn nếu muốn
from ai_processor import call_gemini_api_for_cleanup
from doc_formatter import identify_doc_type as recognize_document_type
from doc_formatter import apply_docx_formatting as format_word_document
from doc_formatter import DOC_TYPE_FORMATTERS, PlaceholderFormatter

app = Flask(__name__)
CORS(app)

# Hàm extract_and_update_signature (giữ nguyên)
def extract_and_update_signature(body_cleaned, data):
    # ... (code giữ nguyên) ...
    extracted_title = None
    extracted_name = None
    pos_match = re.search(r'\[SIGNATURE_POSITION\](.*?)\[/SIGNATURE_POSITION\]', body_cleaned, re.DOTALL)
    if pos_match:
        extracted_title = pos_match.group(1).strip()
        if extracted_title:
            print(f"DEBUG APP: AI extracted signer_title: '{extracted_title}'", flush=True)
            data['signer_title'] = extracted_title
    name_match = re.search(r'\[SIGNATURE_NAME\](.*?)\[/SIGNATURE_NAME\]', body_cleaned, re.DOTALL)
    if name_match:
        extracted_name = name_match.group(1).strip()
        if extracted_name:
             print(f"DEBUG APP: AI extracted signer_name: '{extracted_name}'", flush=True)
             data['signer_name'] = extracted_name
    body_after_tag_removal = re.sub(r'\[SIGNATURE_POSITION\].*?\[/SIGNATURE_POSITION\]', '', body_cleaned, flags=re.DOTALL).strip()
    body_after_tag_removal = re.sub(r'\[SIGNATURE_NAME\].*?\[/SIGNATURE_NAME\]', '', body_after_tag_removal, flags=re.DOTALL).strip()
    return body_after_tag_removal, data

@app.route("/generate", methods=["POST"])
def generate_docx_route():
    print("--- ROUTE /generate START ---", flush=True)
    try:
        data = request.get_json(force=True)
        print(f"  DEBUG APP: Received data keys: {list(data.keys())}", flush=True)

        title = data.get("title", "Văn bản")
        body_original = data.get("body", "")
        intended_doc_type = data.get("intended_doc_type", None) # Lấy loại dự định

        if not body_original:
            print("  ERROR APP: Body is empty.", flush=True)
            return jsonify({"error": "Nội dung văn bản không được để trống"}), 400

        print(f"  DEBUG APP: Input Title: '{title}'", flush=True)
        print(f"  DEBUG APP: Input Intended Type: '{intended_doc_type}'", flush=True)
        # print(f"  DEBUG APP: Input Body (first 100 chars): '{body_original[:100]}...'", flush=True) # Có thể bỏ comment nếu cần xem body

        # Nhận diện loại VB trước khi gọi AI (có thể hữu ích)
        recognized_doc_type = recognize_document_type(title, body_original)
        print(f"  DEBUG APP: Recognized Type BEFORE AI: '{recognized_doc_type}'", flush=True)

        use_ai_cleanup = data.get("use_ai", True) # Kiểm tra cờ use_ai
        body_cleaned = body_original

        if use_ai_cleanup:
            print("  DEBUG APP: Calling AI for cleanup...", flush=True)
            body_cleaned = call_gemini_api_for_cleanup(body_original)
            print(f"  DEBUG APP: Body after AI (first 100 chars): '{body_cleaned[:100]}...'", flush=True)
            # Xử lý tag chữ ký sau AI
            body_cleaned, data = extract_and_update_signature(body_cleaned, data)
            print(f"  DEBUG APP: Body after Tag Removal (first 100 chars): '{body_cleaned[:100]}...'", flush=True)
        else:
            print("  DEBUG APP: Skipping AI cleanup.", flush=True)

        # Cập nhật data['body'] cuối cùng
        data['body'] = body_cleaned
        # Cập nhật lại title nếu AI có thể đã sửa đổi (tuỳ logic mong muốn)
        # data['title'] = ... # Cần xem xét lại logic lấy title sau AI nếu cần

        # Gọi hàm định dạng chính
        print(f"  DEBUG APP: Calling format_word_document with recognized='{recognized_doc_type}', intended='{intended_doc_type}'...", flush=True)
        document, doc_type_for_filename = format_word_document(data, recognized_doc_type, intended_doc_type)
        print(f"  DEBUG APP: format_word_document returned. Doc type for filename: '{doc_type_for_filename}'", flush=True)

        # Lưu và gửi file (giữ nguyên)
        output_stream = io.BytesIO()
        document.save(output_stream)
        output_stream.seek(0)
        safe_title = re.sub(r'[\\/*?:"<>|]', "", title).replace(" ", "_")
        filename = f"{safe_title}_{doc_type_for_filename}_{uuid.uuid4().hex[:6]}.docx"
        print(f"  DEBUG APP: Saving file as: {filename}", flush=True)

        print("--- ROUTE /generate END (Success) ---", flush=True)
        return send_file(
            output_stream, as_attachment=True, download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print(f"!!!!!!!! ERROR in /generate route !!!!!!!!!!", flush=True)
        error_trace = traceback.format_exc()
        print(error_trace, flush=True)
        print("--- ROUTE /generate END (Error) ---", flush=True)
        return jsonify({"error": f"Lỗi server nghiêm trọng: {str(e)}"}), 500

# Route "/" và phần if __name__ == "__main__": giữ nguyên
@app.route("/", methods=["GET"])
def home():
    supported_types = sorted([k for k, v in DOC_TYPE_FORMATTERS.items() if not isinstance(v, PlaceholderFormatter)])
    return jsonify({
        "message": "API tạo văn bản Word chuẩn VN (v5.1 - Debug Logging) sẵn sàng tại /generate",
        "supported_document_types": supported_types
     })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    print(f"--- Starting Flask App on port {port} ---", flush=True)
    app.run(host="0.0.0.0", port=port, debug=False) # Luôn chạy debug=False trên production
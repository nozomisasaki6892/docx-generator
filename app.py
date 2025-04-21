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
from doc_formatter import DOC_TYPE_FORMATTERS # Import để dùng trong route home

app = Flask(__name__)
CORS(app)

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
        if use_ai_cleanup:
            body_cleaned = call_gemini_api_for_cleanup(body_original)
            print("Nội dung đã qua AI làm sạch.")
        else:
            body_cleaned = body_original
            print("Bỏ qua bước AI làm sạch.")
        data['body'] = body_cleaned

        document, doc_type_for_filename = format_word_document(data, recognized_doc_type, intended_doc_type)

        output_stream = io.BytesIO()
        document.save(output_stream)
        output_stream.seek(0)
        safe_title = re.sub(r'[\\/*?:"<>|]', "", title).replace(" ", "_")
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
    # Lấy danh sách các key từ DOC_TYPE_FORMATTERS đã import
    # Lọc bỏ những key có giá trị là None hoặc PlaceholderFormatter
    supported_types = [k for k, v in DOC_TYPE_FORMATTERS.items() if hasattr(v, 'format')]
    return jsonify({
        "message": "API tạo văn bản Word chuẩn VN (v4.1 - Full) sẵn sàng tại /generate",
        "supported_document_types": sorted(supported_types)
     })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
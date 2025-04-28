# app.py (Kiến trúc mới)
import io
import os
import re
import uuid
import traceback
import time
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

# Import config mới
from config import * # Import tất cả để dễ dùng hằng số
# Import hàm AI mới và hàm tạo doc mới
from ai_processor import call_gemini_for_formatted_body, PROMPT_MAP # Import PROMPT_MAP để kiểm tra
from doc_formatter import create_formatted_document # Import hàm mới

app = Flask(__name__)
CORS(app)

# Bỏ hàm extract_and_update_signature

# Hàm helper để chuẩn bị dữ liệu đầu vào
def prepare_data_for_formatting(input_json, intended_doc_type):
    print("  DEBUG APP: Preparing data dictionary...", flush=True)
    data = {}
    current_time = time.localtime()

    # --- Thông tin cơ bản ---
    user_title = input_json.get("title", "Văn bản")
    user_body = input_json.get("body", "")
    data['title'] = user_title # Giữ title gốc cho trích yếu
    data['user_input_data'] = f"Tiêu đề/Yêu cầu: {user_title}\nNội dung chi tiết:\n{user_body}"

    # --- Thông tin Header ---
    # Lấy từ input JSON hoặc đặt giá trị mặc định phù hợp
    data['issuing_org'] = input_json.get("issuing_org", "CHƯA CÓ TÊN CƠ QUAN")
    data['issuing_org_parent'] = input_json.get("issuing_org_parent") # Có thể là None

    # Tách Số và Ký hiệu từ input (cần frontend gửi đúng định dạng)
    doc_number_full = input_json.get("doc_number", f"...") # Ví dụ: "Số: 123/2025/NĐ-CP"
    doc_number_only = "..."
    doc_symbol = "..."
    try:
        parts = doc_number_full.split('/')
        if len(parts) > 1:
            match_num = re.search(r'(\d+)$', parts[0])
            doc_number_only = match_num.group(1) if match_num else "..."
            doc_symbol = parts[1] # Lấy phần sau dấu / đầu tiên
    except Exception:
        print("  WARNING APP: Could not parse doc_number_only and doc_symbol from input.", flush=True)

    data['doc_number_only'] = input_json.get("doc_number_only", doc_number_only)
    data['doc_symbol'] = input_json.get("doc_symbol", doc_symbol)

    # Ngày tháng, địa danh
    data['issuing_location'] = input_json.get("issuing_location", "Hà Nội")
    try:
        data['issuing_day'] = int(input_json.get("issuing_day", current_time.tm_mday))
        data['issuing_month'] = int(input_json.get("issuing_month", current_time.tm_mon))
        data['issuing_year'] = int(input_json.get("issuing_year", current_time.tm_year))
    except (ValueError, TypeError):
        print("  WARNING APP: Invalid date input, using current date.", flush=True)
        data['issuing_day'] = current_time.tm_mday
        data['issuing_month'] = current_time.tm_mon
        data['issuing_year'] = current_time.tm_year

    # --- Thông tin Chữ ký ---
    # Lấy từ input JSON (frontend cần cung cấp các trường này)
    data['authority_signer'] = input_json.get("authority_signer") # VD: "TM. CHÍNH PHỦ"
    data['signer_title'] = input_json.get("signer_title", "[CHỨC VỤ]")
    data['signer_name'] = input_json.get("signer_name", "[Họ tên]")

    # --- Thông tin Nơi nhận ---
    recipients_input = input_json.get("recipients", "- Lưu: VT.") # Lấy từ input
    if isinstance(recipients_input, str):
         data['recipients'] = [r.strip() for r in recipients_input.split('\n') if r.strip()]
    elif isinstance(recipients_input, list):
         data['recipients'] = [str(item) for item in recipients_input] # Đảm bảo là list of strings
    else:
         data['recipients'] = ["- Lưu: VT."]

    # --- Thông tin khác ---
    data['doc_type_label'] = intended_doc_type # Sẽ được map trong doc_formatter

    print(f"  DEBUG APP: Prepared data keys: {list(data.keys())}", flush=True)
    return data


@app.route("/generate", methods=["POST"])
def generate_docx_route():
    print("--- ROUTE /generate START (New Arch v2) ---", flush=True)
    try:
        input_json = request.get_json(force=True)
        print(f"  DEBUG APP: Received raw JSON keys: {list(input_json.keys())}", flush=True)

        # 1. Lấy thông tin cơ bản và kiểm tra
        user_title = input_json.get("title", "Văn bản")
        user_body = input_json.get("body", "")
        intended_doc_type = input_json.get("intended_doc_type") # Loại VB người dùng chọn
        use_ai = input_json.get("use_ai", True) # Mặc định là dùng AI

        if not intended_doc_type:
             print("  ERROR APP: Intended document type is required.", flush=True)
             return jsonify({"error": "Vui lòng chọn loại văn bản muốn tạo"}), 400
        if intended_doc_type not in PROMPT_MAP:
             print(f"  ERROR APP: Unsupported document type '{intended_doc_type}'. No AI prompt defined.", flush=True)
             return jsonify({"error": f"Loại văn bản '{intended_doc_type}' chưa được hỗ trợ hoặc chưa có prompt AI."}), 400
        # Có thể bỏ qua kiểm tra user_body nếu AI có thể tự sinh từ title/yêu cầu
        # if not user_body:
        #     print("  ERROR APP: User body input is empty.", flush=True)
        #     return jsonify({"error": "Nội dung yêu cầu không được để trống"}), 400

        print(f"  DEBUG APP: User Title/Subject: '{user_title}'", flush=True)
        print(f"  DEBUG APP: User Intended Doc Type: '{intended_doc_type}'", flush=True)
        print(f"  DEBUG APP: Use AI flag: {use_ai}", flush=True)

        # 2. Chuẩn bị dictionary `data`
        data = prepare_data_for_formatting(input_json, intended_doc_type)

        # 3. Gọi AI để tạo phần thân đã định dạng (nếu use_ai=True)
        ai_formatted_body = user_body # Mặc định dùng body gốc nếu không dùng AI
        if use_ai:
            print(f"  DEBUG APP: Calling AI for formatted body (type: {intended_doc_type})...", flush=True)
            # Tạo input data cho prompt từ title và body gốc
            user_input_for_ai = f"Tiêu đề/Yêu cầu: {user_title}\nNội dung chi tiết do người dùng cung cấp:\n{user_body}"
            data['user_input_data'] = user_input_for_ai # Cập nhật lại nếu cần thiết

            ai_formatted_body = call_gemini_for_formatted_body(
                user_input_data=data['user_input_data'],
                intended_doc_type=intended_doc_type
            )
            # Kiểm tra nếu AI trả về lỗi
            if "[Lỗi AI]" in ai_formatted_body or "[Lỗi Prompt]" in ai_formatted_body:
                 print(f"  ERROR APP: AI processing failed. Error: {ai_formatted_body}", flush=True)
                 return jsonify({"error": f"Lỗi từ AI: {ai_formatted_body}"}), 500
            print(f"  DEBUG APP: Received AI body (first 100 chars): '{ai_formatted_body[:100]}...'", flush=True)
        else:
            print("  DEBUG APP: Skipping AI call. Using raw user body.", flush=True)

        # Cập nhật data['body'] cuối cùng để truyền cho doc_formatter
        data['body'] = ai_formatted_body

        # 4. Gọi hàm tạo file Word mới
        print(f"  DEBUG APP: Calling create_formatted_document (type: {intended_doc_type})...", flush=True)
        document, final_doc_type_name = create_formatted_document(data, intended_doc_type)
        print(f"  DEBUG APP: create_formatted_document returned. Final type name: '{final_doc_type_name}'", flush=True)

        # Kiểm tra xem có trả về document lỗi không
        if final_doc_type_name.startswith("Error_"):
             print(f"  ERROR APP: create_formatted_document indicated an error.", flush=True)
             # Có thể trả về lỗi hoặc gửi file chứa lỗi tùy ý
             # return jsonify({"error": f"Lỗi tạo file Word cho loại {intended_doc_type}"}), 500

        # 5. Lưu và gửi file
        output_stream = io.BytesIO()
        document.save(output_stream)
        output_stream.seek(0)
        safe_title = re.sub(r'[\\/*?:"<>|]', "", user_title).replace(" ", "_").replace("/", "-")
        filename = f"{safe_title[:50]}_{final_doc_type_name}_{uuid.uuid4().hex[:6]}.docx"
        print(f"  DEBUG APP: Saving and sending file: {filename}", flush=True)

        print("--- ROUTE /generate END (New Arch v2 - Success) ---", flush=True)
        return send_file(
            output_stream,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print(f"!!!!!!!! ERROR in /generate route (New Arch v2) !!!!!!!!!!", flush=True)
        error_trace = traceback.format_exc()
        print(error_trace, flush=True)
        print("--- ROUTE /generate END (New Arch v2 - Error) ---", flush=True)
        return jsonify({"error": f"Lỗi server nghiêm trọng (New Arch v2): {str(e)}"}), 500

# Route "/" và phần if __name__ == "__main__": giữ nguyên
@app.route("/", methods=["GET"])
def home():
     supported_prompts = list(PROMPT_MAP.keys())
     return jsonify({
        "message": "API tạo văn bản Word chuẩn VN (v6.0 - New Arch) sẵn sàng tại /generate",
        "supported_document_types_with_prompts": sorted(supported_prompts)
     })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    print(f"--- Starting Flask App (New Arch v2) on port {port} ---", flush=True)
    app.run(host="0.0.0.0", port=port, debug=False)
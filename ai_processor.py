# ai_processor.py (Kiến trúc mới)
import time
import requests
import re
import traceback
# Import config mới
from config import (
    GEMINI_API_KEY, GEMINI_API_URL,
    MAX_AI_INPUT_LENGTH, AI_RETRY_DELAY,
    # Import các prompt template (Cần thêm các loại khác vào config.py)
    AI_PROMPT_TEMPLATE_NGHI_DINH
    # Ví dụ:
    # AI_PROMPT_TEMPLATE_CONG_VAN,
    # AI_PROMPT_TEMPLATE_QUYET_DINH,
)

# --- Mapping loại văn bản dự định sang template prompt tương ứng ---
# !!! Cần xây dựng đầy đủ map này khi có đủ prompt trong config.py !!!
PROMPT_MAP = {
    "NghiDinh": AI_PROMPT_TEMPLATE_NGHI_DINH,
    "NghiDinhQPPL": AI_PROMPT_TEMPLATE_NGHI_DINH, # Dùng chung prompt NĐ
    # Thêm các loại khác ở đây
    # "CongVan": AI_PROMPT_TEMPLATE_CONG_VAN,
    # "QuyetDinh": AI_PROMPT_TEMPLATE_QUYET_DINH,
}

# --- Hàm gọi AI để tạo thân văn bản đã định dạng ---
def call_gemini_for_formatted_body(user_input_data, intended_doc_type):
    """
    Gọi Gemini API để tạo phần THÂN VĂN BẢN đã được ĐỊNH DẠNG SẴN
    dựa trên prompt chi tiết tương ứng với loại văn bản.
    """
    print(f"--- AI Processor START: Generating formatted body for type '{intended_doc_type}' ---", flush=True)
    if not GEMINI_API_KEY or GEMINI_API_KEY == "YOUR_API_KEY_HERE_IF_NO_ENV":
        print("  WARNING AI Processor: GEMINI_API_KEY not set. Returning raw input as body.", flush=True)
        return f"[Lỗi API Key] Phần thân văn bản chưa được AI định dạng:\n{user_input_data}"

    # 1. Chọn đúng prompt template
    prompt_template = PROMPT_MAP.get(intended_doc_type)
    if not prompt_template:
        error_message = f"[Lỗi Prompt] Không tìm thấy prompt template cho loại văn bản '{intended_doc_type}'. Vui lòng định nghĩa trong config.py và PROMPT_MAP trong ai_processor.py."
        print(f"  ERROR AI Processor: {error_message}", flush=True)
        return f"{error_message}\nNội dung gốc:\n{user_input_data}"

    # 2. Chuẩn bị prompt cuối cùng
    try:
        final_prompt = prompt_template.format(user_input_data=user_input_data)
    except KeyError as key_error:
         error_message = f"[Lỗi Prompt Format] Prompt template cho '{intended_doc_type}' thiếu placeholder {key_error}. Vui lòng kiểm tra lại config.py."
         print(f"  ERROR AI Processor: {error_message}", flush=True)
         return f"{error_message}\nNội dung gốc:\n{user_input_data}"

    print(f"  DEBUG AI Processor: Using prompt template for '{intended_doc_type}'. Prompt length: {len(final_prompt)}", flush=True)

    # 3. Xử lý input dài (Tạm thời gửi cả prompt lớn, cần xem xét lại nếu prompt quá dài)
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": final_prompt}]}],
        "generationConfig": {"temperature": 0.6, "maxOutputTokens": 8192}
    }

    # 4. Gọi API và xử lý kết quả
    retries = 3
    for attempt in range(retries):
        print(f"  DEBUG AI Processor: Calling Gemini API (Attempt {attempt + 1}/{retries})...", flush=True)
        try:
            response = requests.post(f"{GEMINI_API_URL}?key={GEMINI_API_KEY}", headers=headers, json=payload, timeout=180) # Tăng timeout
            response.raise_for_status() # Ném lỗi nếu status code là 4xx hoặc 5xx
            result = response.json()

            # Kiểm tra kỹ cấu trúc response trả về từ Gemini
            candidate = result.get('candidates', [{}])[0]
            content = candidate.get('content', {})
            parts = content.get('parts', [{}])
            formatted_body = parts[0].get('text', '')

            if formatted_body:
                # Hậu xử lý cơ bản (xóa ``` nếu có)
                formatted_body = re.sub(r'^```(python|text|markdown)?\n', '', formatted_body, flags=re.IGNORECASE)
                formatted_body = re.sub(r'\n```$', '', formatted_body)
                print(f"  DEBUG AI Processor: Received formatted body from Gemini. Length: {len(formatted_body)}", flush=True)
                print("--- AI Processor END (Success) ---", flush=True)
                return formatted_body.strip()
            # Xử lý trường hợp bị chặn
            elif candidate.get('finishReason') == 'SAFETY':
                 safety_ratings = candidate.get('safetyRatings', [])
                 block_reason_detail = "; ".join([f"{r.get('category')}: {r.get('probability')}" for r in safety_ratings])
                 error_message = f"[Lỗi AI] Nội dung bị chặn vì lý do an toàn: {block_reason_detail}. Hãy thử lại với nội dung khác."
                 print(f"  ERROR AI Processor: {error_message}", flush=True)
                 return error_message
            # Xử lý các trường hợp trả về không có nội dung khác
            else:
                error_message = f"[Lỗi AI] API trả về không có nội dung hoặc định dạng không mong muốn. Response: {result}"
                print(f"  ERROR AI Processor: {error_message}", flush=True)
                if attempt < retries - 1:
                     print(f"  Retrying after {AI_RETRY_DELAY} seconds...", flush=True)
                     time.sleep(AI_RETRY_DELAY)
                else:
                     return error_message # Trả lỗi sau lần thử cuối

        except requests.exceptions.RequestException as e:
            error_message = f"[Lỗi Kết Nối AI] Không thể kết nối đến Gemini API: {e}"
            print(f"  ERROR AI Processor: API Call failed: {e}", flush=True)
            if attempt < retries - 1:
                 print(f"  Retrying after {AI_RETRY_DELAY} seconds...", flush=True)
                 time.sleep(AI_RETRY_DELAY)
            else:
                 print("--- AI Processor END (Failed) ---", flush=True)
                 return error_message # Trả lỗi sau lần thử cuối
        except Exception as e:
            error_message = f"[Lỗi AI Processor] Lỗi không xác định: {e}"
            print(f"!!!!!!!! ERROR in call_gemini_for_formatted_body !!!!!!!!!!", flush=True)
            print(traceback.format_exc(), flush=True)
            # Không retry với lỗi không xác định, trả lỗi ngay
            print("--- AI Processor END (Error) ---", flush=True)
            return error_message

    # Trả về lỗi nếu hết số lần thử mà không thành công
    return "[Lỗi AI] Không thể tạo nội dung body sau nhiều lần thử."
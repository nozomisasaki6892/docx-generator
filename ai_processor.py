# ai_processor.py
import time
import requests
import re
from config import GEMINI_API_KEY, GEMINI_API_URL, AI_PROMPT_TEMPLATE, MAX_AI_INPUT_LENGTH, AI_RETRY_DELAY

def call_gemini_api_for_cleanup(text_to_clean):
    """Gọi Gemini API để làm sạch nội dung văn bản."""
    if not GEMINI_API_KEY or GEMINI_API_KEY == "YOUR_API_KEY_HERE_IF_NO_ENV":
        print("WARNING: GEMINI_API_KEY is not set. Skipping AI cleanup.")
        return text_to_clean

    headers = {"Content-Type": "application/json"}
    parts = [text_to_clean[i:i+MAX_AI_INPUT_LENGTH] for i in range(0, len(text_to_clean), MAX_AI_INPUT_LENGTH)]
    cleaned_parts = []
    print(f"Chia văn bản thành {len(parts)} phần để xử lý AI.")

    for i, part in enumerate(parts):
        print(f"Đang xử lý làm sạch phần {i+1}/{len(parts)} bằng Gemini...")
        prompt = AI_PROMPT_TEMPLATE.format(text_input=part)
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.5, "maxOutputTokens": 8192}
        }
        try:
            response = requests.post(f"{GEMINI_API_URL}?key={GEMINI_API_KEY}", headers=headers, json=payload, timeout=120)
            response.raise_for_status()
            result = response.json()
            if 'candidates' in result and len(result['candidates']) > 0 and 'content' in result['candidates'][0] and 'parts' in result['candidates'][0]['content'] and len(result['candidates'][0]['content']['parts']) > 0:
                cleaned_text = result['candidates'][0]['content']['parts'][0]['text']
                cleaned_text = re.sub(r'^```.*?\n', '', cleaned_text, flags=re.MULTILINE)
                cleaned_text = re.sub(r'^>', '', cleaned_text, flags=re.MULTILINE)
                cleaned_parts.append(cleaned_text.strip())
                print(f"Làm sạch xong phần {i+1}.")
            else:
                print(f"Lỗi: Không tìm thấy nội dung trả về từ Gemini cho phần {i+1}. Response: {result}")
                cleaned_parts.append(part)
            if i < len(parts) - 1:
                print(f"Nghỉ {AI_RETRY_DELAY} giây...")
                time.sleep(AI_RETRY_DELAY)
        except requests.exceptions.RequestException as e:
            print(f"Lỗi gọi Gemini API cho phần {i+1}: {e}")
            cleaned_parts.append(part)
        except Exception as e:
            print(f"Lỗi không xác định khi xử lý Gemini response phần {i+1}: {e}")
            cleaned_parts.append(part)

    print("Ghép các phần đã làm sạch...")
    full_cleaned_text = "\n".join(cleaned_parts)
    return full_cleaned_text.strip()

# Có thể thêm các hàm xử lý AI khác ở đây (ví dụ: nhận diện loại VB bằng AI)
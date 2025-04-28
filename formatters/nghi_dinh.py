# formatters/nghi_dinh.py (PHIÊN BẢN SIÊU TỐI GIẢN DEBUG)
# Không import bất cứ thứ gì khác ngoài Document nếu cần (nhưng hàm format nhận document rồi)
import traceback

def format(document, data):
    print("--- RUNNING DEBUG MINIMAL nghi_dinh.py ---", flush=True)
    try:
        # Thử thêm một dòng text duy nhất, không định dạng gì cả
        document.add_paragraph("<<<<< NẾU BẠN THẤY DÒNG NÀY, NGHI_DINH.PY ĐÃ CHẠY! >>>>>")
        print("--- DEBUG MINIMAL: Added test paragraph. ---", flush=True)
    except Exception as e:
        print(f"!!! ERROR in Minimal Debug nghi_dinh.py: {e}", flush=True)
        print(traceback.format_exc(), flush=True)
        try:
            # Cố gắng ghi lỗi vào doc
            document.add_paragraph(f"!!! ERROR in Minimal Debug: {e} !!!")
        except:
            pass # Bỏ qua nếu không ghi được lỗi
    print("--- FINISHED DEBUG MINIMAL nghi_dinh.py ---", flush=True)
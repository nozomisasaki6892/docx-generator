from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import uuid
import io

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate_docx():
    data = request.get_json()
    title = data.get("title", "CÔNG VĂN")
    body = data.get("body", "Kính gửi đơn vị liên quan,\nNội dung công văn sẽ được cập nhật tại đây.")

    document = Document()

    # Thiết lập lề văn bản (chuẩn A4, hành chính)
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(3.5)
    section.right_margin = Cm(2.0)

    # Style chung
    style = document.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(13)

    # Quốc hiệu
    p1 = document.add_paragraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p1.runs[0]
    run1.bold = True

    # Tiêu ngữ
    p2 = document.add_paragraph("Độc lập - Tự do - Hạnh phúc")
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.runs[0]
    run2.bold = True

    # Dòng kẻ ngang
    document.add_paragraph("_____________________________").alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Tiêu đề công văn
    document.add_paragraph(title).alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Nội dung chính
    for line in body.split("\n"):
        if line.strip():
            document.add_paragraph(line.strip())

    # Xuất file
    output_stream = io.BytesIO()
    document.save(output_stream)
    output_stream.seek(0)

    filename = f"congvan_{uuid.uuid4().hex[:8]}.docx"
    return send_file(output_stream, as_attachment=True, download_name=filename)

@app.route("/", methods=["GET"])
def home():
    return jsonify({"message": "API công văn sẵn sàng tại /generate"})

if __name__ == "__main__":
    app.run(debug=True)

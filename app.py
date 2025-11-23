import os
import threading
from flask import Flask, render_template, request, send_from_directory, jsonify
from utils.bank_shinhan_extract import run_ocr_pipeline

app = Flask(__name__)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
RESULT_DIR = os.path.join(os.path.dirname(__file__), "results")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

# 진행 상태 전역 저장용
progress = {}
upload_names = {}

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    """
    1️⃣ 파일 업로드 및 비동기 OCR 실행
    """
    file = request.files["pdf_file"]

    if not file:
        return "파일을 업로드하세요.", 400

    file_id = str(os.urandom(8).hex())

    # 원본 파일명 저장 (확장자 제외)
    original_name = os.path.splitext(file.filename)[0]
    upload_names[file_id] = original_name

    # 파일 저장
    saved_name = f"{file_id}_{file.filename}"
    file_path = os.path.join(UPLOAD_DIR, saved_name)

    result_path = os.path.join(RESULT_DIR, f"{file_id}.xlsx")
    file.save(file_path)

    progress[file_id] = 0

    # 백그라운드에서 OCR 실행
    thread = threading.Thread(target=process_pdf, args=(file_id, file_path, result_path))
    thread.start()

    return render_template("progress.html", task_id=file_id)


def process_pdf(file_id, pdf_path, result_path):
    """실제 OCR 처리 로직"""
    try:
        # 진행률 업데이트 콜백 정의
        def update_progress(task_id, percent):
            # 간단히 전역 dict 업데이트
            progress[task_id] = int(percent)

        # OCR 파이프라인 실행 + 진행률 콜백 전달
        run_ocr_pipeline(
            pdf_path=pdf_path,
            output_path=result_path,
            task_id=file_id,
            progress_callback=update_progress,
        )
    except Exception as e:
        print(f"[ERROR] {e}")
        progress[file_id] = -1


@app.route("/progress/<task_id>")
def check_progress(task_id):
    """AJAX 폴링용 진행률 API"""
    # 캐시 방지 헤더 추가
    from flask import make_response
    value = progress.get(task_id, 0)
    resp = make_response(jsonify({"progress": value}))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

@app.route("/result/<task_id>")
def show_result(task_id):
    return render_template("result.html", task_id=task_id)

@app.route("/download/<task_id>")
def download_result(task_id):
    """결과 Excel 다운로드"""
    stored_filename = f"{task_id}.xlsx"

    original_name = upload_names.get(task_id, f"result_{task_id}")
    download_name = f"{original_name}_결과.xlsx"

    return send_from_directory(RESULT_DIR, stored_filename, as_attachment=True, download_name=download_name)


if __name__ == "__main__":
    app.run(debug=True, port=5000)

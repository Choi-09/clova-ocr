import importlib
import os
import threading
from flask import Flask, render_template, request, send_from_directory, jsonify

app = Flask(__name__)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
RESULT_DIR = os.path.join(os.path.dirname(__file__), "results")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

# 진행 상태 전역 저장용
progress = {}
upload_names = {}

# 1) 파일명에서 은행 코드 추출
# TODO DB에서 가져오기
INSTITUTION_PATTERNS = {
    "shinhan": ["신한", "shinhan"],
    "kb": ["국민", "kookmin", "kb"],
    "nh": ["농협", "nh"],
    "woori": ["우리", "woori"],
    "hana": ["하나", "hana"],
    # 카드 전용 회사 (은행 없는 카드사)도 여기 추가 가능: "lotte", "hyundai" 등
}
PRODUCT_PATTERNS = {
    "bank": ["은행", "bank"],
    "card": ["카드", "card"],
    # 필요하면 "loan": ["대출"], "insurance": ["보험"] 같은 것도 확장 가능
}


def detect_pipeline_code_from_filename(filename: str) -> str | None:
    """
    파일명(확장자 제외)에서
      1) 기관명 (shinhan, kb, nh, ...)
      2) 상품 타입 (bank, card, ...)
    를 각각 찾은 뒤 "{product}_{institution}" 형태의 pipeline code 를 만든다.
    예:
      '신한은행.pdf'  → 'bank_shinhan'
      '신한카드.pdf'  → 'card_shinhan'
    """
    if not filename:
        return None

    base = os.path.splitext(filename)[0]
    lower = base.lower()

    institution = None
    product = None

    # 1) 기관명 찾기
    for inst_code, patterns in INSTITUTION_PATTERNS.items():
        for p in patterns:
            if (p.isascii() and p in lower) or (not p.isascii() and p in base):
                institution = inst_code
                break
        if institution:
            break

    # 2) 상품 타입 찾기 (은행 / 카드 등)
    for prod_code, patterns in PRODUCT_PATTERNS.items():
        for p in patterns:
            if (p.isascii() and p in lower) or (not p.isascii() and p in base):
                product = prod_code
                break
        if product:
            break

    if not institution or not product:
        # 하나라도 못 찾으면 None 반환 (별도 에러 처리)
        return None

    return f"{product}_{institution}"

# 2) 금융사별 OCR 파이프라인 동적 로드
def get_ocr_pipeline(pipeline_code: str):
    """
    금융사 코드에 따라 해당 모듈의 run_ocr_pipeline 을 동적으로 import.
    예:
      'shinhan' → utils.bank_shinhan_extract
      'kb'      → utils.bank_kb_extract
    """
    module_name = f"utils.{pipeline_code}_extract"
    try:
        module = importlib.import_module(module_name)
        return module.run_ocr_pipeline
    except ModuleNotFoundError:
        raise ValueError(f"현재 서비스 준비중입니다.: {pipeline_code}")

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

    # 파일명에서 은행 코드 추론
    pipeline_code  = detect_pipeline_code_from_filename(file.filename)
    print("detected pipeline_code:", pipeline_code)

    if not pipeline_code:
        # 은행명을 못 찾았을 때의 에러 메시지는 필요에 따라 조정
        return "금융사를 식별할 수 없습니다. 파일명에 은행명을 포함해 주세요.(예: 신한은행_거래내역.pdf, shinhancard_거래내역.pdf 등)", 400

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
    thread = threading.Thread(target=process_pdf, args=(pipeline_code, file_id, file_path, result_path))
    thread.start()

    return render_template("progress.html", task_id=file_id)


def process_pdf(pipeline_code: str, file_id: str, pdf_path: str, result_path: str):
    """실제 OCR 처리 로직"""
    try:
        # 금융사별 OCR 파이프라인 함수 동적 로드
        ocr_pipeline = get_ocr_pipeline(pipeline_code)

        # 진행률 업데이트 콜백 정의
        def update_progress(task_id, percent):
            # 간단히 전역 dict 업데이트
            progress[task_id] = int(percent)

        # OCR 파이프라인 실행 + 진행률 콜백 전달
        ocr_pipeline(
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
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )

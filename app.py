import os
import threading
from flask import Flask, render_template, request, send_file, jsonify
from utils.bank_shinhan_process import process_pdf_and_ocr

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULT_FOLDER'] = 'results'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULT_FOLDER'], exist_ok=True)

progress_status = {}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'pdf_file' not in request.files:
        return "파일이 없습니다", 400
    file = request.files['pdf_file']

    file = request.files['pdf_file']
    if file.filename == '':
        return "파일이 선택되지 않았습니다", 400

    filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    progress_status[filename] = 0
    thread = threading.Thread(target=process_task, args=(filename, filepath))
    thread.start()

    return render_template('progress.html', filename=filename)

def process_task(filename, filepath):
    output_excel = os.path.splitext(filename)[0] + '_결과.xlsx'
    output_excel_path = os.path.join(app.config['RESULT_FOLDER'], output_excel)

    def update_progress(percent):
        progress_status[filename] = percent

    wb = process_pdf_and_ocr(filepath, update_progress_callback=update_progress)
    wb.save(output_excel_path)

    progress_status[filename] = 100  # 완료 상태

@app.route('/progress/<filename>')
def get_progress(filename):
    status = progress_status.get(filename, 0)
    return jsonify({'progress': status})

@app.route('/result/<filename>')
def result(filename):
    return render_template('result.html', filename=filename)

@app.route('/download/<filename>')
def download_file(filename):
    output_excel_path = os.path.join(app.config['RESULT_FOLDER'], filename)

    if not os.path.exists(output_excel_path):
        return f"File not found: {filename}", 404

    return send_file(
        output_excel_path,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=True)

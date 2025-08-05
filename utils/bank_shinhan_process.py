# import csv
# import json
# import time
# import uuid
#
# import requests
# import re
#
# api_url = 
# secret_key = 
# image_file_path = 'C:\\Study\\law-vo\\'
# image_file_name = '신한은행-거래내역'
# file_format = '.pdf'
# image_file = image_file_path + image_file_name + file_format
# # OCR 요청
# request_json = {
#     'images': [
#         {
#             'format': 'pdf',
#             'name': 'demo'
#         }
#     ],
#     'requestId': str(uuid.uuid4()),
#     'version': 'V2',
#     'timestamp': int(round(time.time() * 1000))
# }
#
# payload = {'message': json.dumps(request_json).encode('UTF-8')}
# files = [('file', open(image_file,'rb'))]
# headers = {'X-OCR-SECRET': secret_key}
# response = requests.post(api_url, headers=headers, data=payload, files=files)
# result = json.loads(response.text)
# # print("result: ", result)
#
# # 모든 텍스트 추출
# fields = []
# # detail_field = []
# print(json.dumps(result['images'], indent=2, ensure_ascii=False))
#
# for image in result['images']:
#     if image.get('fields'):
#         fields.extend(image['fields'])
#     elif 'title' in image and 'subFields' in image['title']:
#         fields.extend(image['title']['subFields'])
#         print("fields: ", fields)
#     else:
#         pass
#
# basic_info_text = ' '.join(field.get('inferText', '') for field in fields)
# print("basic_info_text: ", basic_info_text)
#
# # 기본 정보 추출 함수
# def extract_basic_info(text):
#     info = {}
#     patterns = {
#         '계좌번호': r'계좌번호[:\s]*([\d\-]+)',
#         '예금주': r'예금주 성명[:\s]*([^\s]+)',
#         '상품명': r'상품명[:\s]*([^\s]+)',
#         '조회기간': r'조회기간[:\s]*([\d\-]+)\s* \s*([\d\-]+)',
#         '업무구분': r'업무구분 1:[:\s]*([\d:가-힣]+)'
#     }
#     for key, pattern in patterns.items():
#         m = re.search(pattern, text)
#         if m:
#             if key == '조회기간':
#                 info[key] = f"{m.group(1)} ~ {m.group(2)}"
#             else:
#                 info[key] = m.group(1)
#         else:
#             info[key] = ''
#     return info
#
# basic_info = extract_basic_info(basic_info_text)
#
# print("basic_info: ", basic_info)
#
# # 거래내역 추출
# def extract_transaction_table_by_position(fields, y_threshold=5):
#     # 1. 필드 정리 (텍스트 + 중심 좌표 추출)
#     items = []
#     for field in fields:
#         text = field.get('inferText', '').strip()
#         vertices = field.get('boundingPoly', {}).get('vertices', [])
#         if not text or not vertices:
#             continue
#         # x = sum(v.get('x', 0) for v in vertices) / len(vertices)
#         x = min(v.get('x', 0) for v in vertices)
#         y = sum(v.get('y', 0) for v in vertices) / len(vertices)
#         items.append({'text': text, 'x': int(x), 'y': int(y)})
#
#     # 2. y 기준 정렬 및 행 그룹화
#     items.sort(key=lambda x: x['y'])
#     rows = []
#     current_row = []
#     prev_y = None
#
#     for item in items:
#         if prev_y is None or abs(item['y'] - prev_y) <= y_threshold:
#             current_row.append(item)
#         else:
#             rows.append(current_row)
#             current_row = [item]
#         prev_y = item['y']
#     if current_row:
#         rows.append(current_row)
#
#     # 3. 올바른 헤더 행 탐색
#     header_row = None
#     header_y = None
#     for row in rows:
#         texts = [item['text'] for item in row]
#         joined_text = ''.join(texts)
#         if all(key in joined_text for key in ['거래일자', '내용', '찾으신금액', '맡기신금액']):
#             header_row = sorted(row, key=lambda x: x['x'])
#             header_y = sum(item['y'] for item in row) / len(row)
#             break
#     if not header_row:
#         raise ValueError("헤더 행을 찾을 수 없습니다.")
#
#     # 4. 헤더 병합 처리 및 위치 설정
#     merged_header = []
#     merged_positions = []
#     i = 0
#     while i < len(header_row):
#         this_text = header_row[i]['text']
#         if i + 1 < len(header_row):
#             next_text = header_row[i + 1]['text']
#             merged = this_text + next_text
#             if merged in ['비고', '잔액']:
#                 merged_header.append(merged)
#                 merged_positions.append((header_row[i]['x'] + header_row[i + 1]['x']) // 2)
#                 i += 2
#                 continue
#         merged_header.append(this_text)
#         merged_positions.append(header_row[i]['x'])
#         i += 1
#
#     # 5. x 위치 기준 경계 계산
#     correction = 10
#     x_ranges = []
#     for i in range(len(merged_positions)):
#         if i == 0:
#             left = -float('inf')
#         else:
#             left = (merged_positions[i - 1] + merged_positions[i]) // 2
#         if i == len(merged_positions) - 1:
#             right = float('inf')
#         else:
#             right = (merged_positions[i] + merged_positions[i + 1]) // 2
#         x_ranges.append((left - correction, right + correction))
#
#     # 6. 헤더 아래 행 필터링 및 정렬
#     aligned_rows = []
#     for row in rows:
#         # 헤더보다 위에 있는 경우 무시
#         avg_y = sum(item['y'] for item in row) / len(row)
#         if avg_y <= header_y:
#             continue
#
#         # 거래일자 패턴 확인
#         row_texts = [item['text'] for item in row]
#         if not any(re.match(r'\d{4}-\d{2}-\d{2}', text) for text in row_texts):
#             continue
#
#         aligned = [' '] * len(merged_positions)
#         for item in row:
#             x = item['x']
#             for idx, (left, right) in enumerate(x_ranges):
#                 if left <= x < right:
#                     aligned[idx] = item['text']
#                     break
#         aligned_rows.append(aligned)
#
#     return merged_header, aligned_rows
#
#
#
# header, rows = extract_transaction_table_by_position(fields)
# print("header:" , header,  "rows: ", rows)
#
# def replace_W_before_number(text):
#     return re.sub(r'W(?=\d)', r'\\', text)
#
# csv_file_name = f'{image_file_name}.csv'
# with open(csv_file_name, mode='w', newline='', encoding='utf-8-sig') as f:
#     writer = csv.writer(f)
#
#     # 기본 정보
#     writer.writerow(['[기본 정보]'])
#     for key, value in basic_info.items():
#         writer.writerow([key, value])
#
#     writer.writerow([])
#     writer.writerow(['[거래 내역]'])
#
#     # 거래 내역
#     writer.writerow(header)
#     for row in rows:
#         new_row = [replace_W_before_number(cell) for cell in row]
#
#         if isinstance(new_row, list):
#             writer.writerow(new_row)
#         elif isinstance(new_row, str):
#             writer.writerow(new_row.split(','))
#         else:
#             print("Unexpected row format:", new_row)


import os
import uuid
import time
import re
import requests
import json
from PyPDF2 import PdfReader, PdfWriter
from openpyxl import Workbook
from dotenv import load_dotenv


load_dotenv() 

# 기본 설정
API_URL = os.getenv("API_URL")
print("API_URL: ", API_URL)
SECRET_KEY = os.getenv("SECRET_KEY")

# PDF -> 단일 페이지 PDF로 분할
def split_pdf_pages(pdf_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    reader = PdfReader(pdf_path)
    paths = []

    base_filename = os.path.splitext(os.path.basename(pdf_path))[0]

    for i, page in enumerate(reader.pages):
        path = os.path.join(output_dir, f"{base_filename}-{i+1}.pdf")
        with open(path, "wb") as f:
            writer = PdfWriter()
            writer.add_page(page)
            writer.write(f)
        paths.append(path)
    return paths

# OCR 요청
def call_ocr_api(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            payload = {
                'message': json.dumps({
                    'images': [{'format': 'pdf', 'name': 'demo'}],
                    'requestId': str(uuid.uuid4()),
                    'version': 'V2',
                    'timestamp': int(time.time() * 1000)
                }).encode('UTF-8')
            }
            files = [('file', file)]
            headers = {'X-OCR-SECRET': SECRET_KEY}
            response = requests.post(API_URL, headers=headers, data=payload, files=files)
            return response.json()
    except Exception as e:
        print(f"[ERROR] OCR API 호출 실패: {e}")
        return None

# 기본 정보 추출
def extract_basic_info(text):
    patterns = {
        '계좌번호': r'계좌번호[:\s]*([\d\-]+)',
        '예금주': r'예금주 성명[:\s]*([^\s]+)',
        '상품명': r'상품명[:\s]*([^\s]+)',
        '조회기간': r'조회기간[:\s]*([\d\-]+)\s* \s*([\d\-]+)',
        '업무구분': r'업무구분 1:[:\s]*([\d:가-힣]+)'
    }
    info = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if key == '조회기간' and match:
            info[key] = f"{match.group(1)} ~ {match.group(2)}"
        else:
            info[key] = match.group(1) if match else ''
    return info

# 거래내역 추출
def extract_transaction_table_by_position(fields, y_threshold=5):
    # 1. 필드 정리 (텍스트 + 중심 좌표 추출)
    items = [
        {
            'text': f.get('inferText', '').strip(),
            'x': min(v.get('x', 0) for v in f.get('boundingPoly', {}).get('vertices', [])),
            'y': sum(v.get('y', 0) for v in f.get('boundingPoly', {}).get('vertices', [])) / 4
        }
        for f in fields if f.get('inferText') and f.get('boundingPoly')
    ]
    items.sort(key=lambda x: x['y'])

    # 2. y 기준 정렬 및 행 그룹화
    rows, row, prev_y = [], [], None
    for item in items:
        if prev_y is None or abs(item['y'] - prev_y) <= y_threshold:
            row.append(item)
        else:
            rows.append(row)
            row = [item]
        prev_y = item['y']
    if row:
        rows.append(row)

    # 3. 헤더 찾기
    for row in rows:
        joined_text = ''.join([item['text'] for item in row])
        if all(key in joined_text for key in ['거래일자', '내용', '찾으신금액', '맡기신금액']):
            header_row = sorted(row, key=lambda x: x['x'])
            header_y = sum(item['y'] for item in row) / len(row)
            break
    else:
        raise ValueError("헤더를 찾을 수 없습니다.")

    # 4. 헤더 병합 처리 및 위치 설정
    merged_header, merged_x = [], []
    i = 0
    while i < len(header_row):
        curr = header_row[i]['text']
        if i + 1 < len(header_row):
            next_text = header_row[i + 1]['text']
            if curr + next_text in ['비고', '잔액']:
                merged_header.append(curr+ next_text)
                merged_x.append((header_row[i]['x'] + header_row[i + 1]['x']) // 2)
                i += 2
                continue
        merged_header.append(curr)
        merged_x.append(header_row[i]['x'])
        i += 1

    # 5. x 위치 기준 경계 계산
    correction = 10
    x_ranges = []
    for i in range(len(merged_x)):
        left = (merged_x[i - 1] + merged_x[i]) // 2 if i > 0 else -float('inf')
        right = (merged_x[i] + merged_x[i + 1]) // 2 if i < len(merged_x) - 1 else float('inf')
        x_ranges.append((left - correction, right + correction))

    # 6. 거래내역 행 정렬
    aligned_rows = []
    for row in rows:
        if sum(i['y'] for i in row) / len(row) <= header_y:
            continue
        if not any(re.match(r'\d{4}-\d{2}-\d{2}', i['text']) for i in row):
            continue
        aligned = [''] * len(merged_x)

        for item in row:
            for idx, (left, right) in enumerate(x_ranges):
                if left <= item['x'] < right:
                    aligned[idx] = item['text']
                    break
        aligned_rows.append(aligned)
    return merged_header, aligned_rows

# 숫자 앞 문자 원화 기호로 바꾸기
def replace_W_before_number(text):
    return re.sub(r'[W\\](?=\d)', '₩', text)

# Excel 저장
def save_to_excel(wb, sheet_name, basic_info, header, rows):
    ws = wb.create_sheet(title=sheet_name)
    ws.append(["[기본 정보]"])
    for key, value in basic_info.items():
        ws.append([key, value])
    ws.append([])
    ws.append([f"[{sheet_name} - 거래 내역]"])
    ws.append(header)
    for row in rows:
        ws.append([replace_W_before_number(cell) for cell in row])

# 실행
def process_pdf_and_ocr(input_pdf_path, update_progress_callback=None):
    temp_dir = os.path.join('temp_split', os.path.splitext(os.path.basename(input_pdf_path))[0])

    pdf_pages = split_pdf_pages(input_pdf_path, temp_dir)
    total_pages = len(pdf_pages)

    wb = Workbook()
    wb.remove(wb.active)

    for idx, page_pdf in enumerate(pdf_pages):
        print(f"📄 처리 중: 페이지 {idx + 1}")
        result = call_ocr_api(page_pdf)
        if not result or 'images' not in result:
            print(f"[⚠️] 페이지 {idx + 1} OCR 실패")
            if update_progress_callback:
                update_progress_callback((idx + 1) / total_pages * 100)
            continue

        images = result['images']
        total_images = len(images)

        fields = []
        for img_idx, image in enumerate(images):
            if image.get('fields'):
                fields.extend(image['fields'])
            elif 'title' in image and 'subFields' in image['title']:
                fields.extend(image['title']['subFields'])

            if update_progress_callback:
                progress = (idx + (img_idx + 1) / total_images) / total_pages
                update_progress_callback(int(progress * 100))

        text = ' '.join(field.get('inferText', '') for field in fields)
        basic_info = extract_basic_info(text)

        try:
            header, rows = extract_transaction_table_by_position(fields)
            save_to_excel(wb, f"페이지 {idx + 1}", basic_info, header, rows)
        except Exception as e:
            print(f"[⚠️] 페이지 {idx + 1} 처리 오류: {e}")


    return wb

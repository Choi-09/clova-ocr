import json
import os
import re
import time
import uuid

import pandas as pd
import requests
from PyPDF2 import PdfReader, PdfWriter
from dotenv import load_dotenv

load_dotenv()

# 기본 설정
API_URL = os.getenv("API_URL")
SECRET_KEY = os.getenv("SECRET_KEY")

# PDF -> 단일 페이지 PDF로 분할
def split_pdf_pages(pdf_path, output_dir):
    """
    PDF 파일을 페이지별로 분할하여 개별 PDF 파일로 저장
    """
    os.makedirs(output_dir, exist_ok=True)
    reader = PdfReader(pdf_path)
    paths = []

    base_filename = os.path.splitext(os.path.basename(pdf_path))[0]

    for i, page in enumerate(reader.pages):
        path = os.path.join(output_dir, f"{base_filename}-{i + 1}.pdf")
        with open(path, "wb") as f:
            writer = PdfWriter()
            writer.add_page(page)
            writer.write(f)
        paths.append(path)
    return paths

# API 콜
def call_ocr_api(pdf_path):
    """
    네이버 CLOVA OCR API를 호출하여 PDF 파일에서 텍스트 추출
    """
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
            response.raise_for_status() # HTTP 오류 발생 시 예외 발생
            return response.json()
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] OCR API 호출 실패: {e}")
        return None
    except Exception as e:
        print(f"[ERROR] OCR API 처리 중 예상치 못한 오류 발생: {e}")
        return None

template_format = {}

# 템플릿 포멧 추출
def get_template_info(result):
    template_name = None
    fields = []
    try:
        images = result.get('images', None)
        if images:
            image = images[0]
            matched_template = image.get('matchedTemplate', {})
            template_name = matched_template.get('name', '')
            fields = image.get('fields', [])

    except Exception as e:
        print(f"[ERROR] 템플릿 정보 추출 실패: {e}")

    template_format['template_name'] = template_name
    template_format['fields'] = fields


# 기본 정보 추출
# 1페이지 판별(fields.name = '기본정보')
def extract_basic_info():
    template_name = template_format.get('template_name')
    fields = template_format.get('fields', [])

    if template_name == 'bank_shinhan_1' and fields:
        # 첫 번째 필드가 '기본정보'인지 확인
        first_field = fields[0]
        if first_field.get('name') == '기본정보':
            # TODO first_field.get('subFields') 에서 'inferText' 내용 추출,
            #  'boundingPoly'의 좌표 추출해서 dataFrame으로 정렬=> df리턴
            basic_info = first_field.get('inferText')
            return basic_info

    #  template_name !== 'bank_shinhan_1'인 경우
    return None

def parse_basic_info(basic_info_text):
    """
    OCR 기본정보 텍스트를 각 항목별로 분리
    """
    if not basic_info_text:
        return {}

    # 한 줄로 합치기
    text = basic_info_text.replace('\n', ' ')

    # 추출할 항목 리스트
    keys = ['계좌번호', '조회기간', '예금주 성명', '상품명']

    result = {}
    for i, key in enumerate(keys):
        if i < len(keys) - 1:
            next_key = keys[i + 1]
            pattern = rf'{key}\s+(.*?)(?=\s{next_key})'
        else:
            pattern = rf'{key}\s+(.*)'

        match = re.search(pattern, text)
        if match:
            value = match.group(1).strip()
            # 조회기간 처리: 두 날짜 사이에 ~ 삽입
            if key == '조회기간':
                dates = value.split()
                if len(dates) == 2:
                    value = f"{dates[0]} ~ {dates[1]}"
            result[key] = value
    return result

# 거래일자 컬럼기준 y 좌표 판별
def rows_by_reference_y_with_text(target_fields, reference_col_index=0, y_threshold=5):
    """
    거래일자 컬럼 기준으로 y 좌표를 맞추고, 각 셀은 첫 번째 \n 이후 텍스트를 사용
    """
    # 1. 각 컬럼별 (y, text) 리스트 생성
    cols = []
    for f in target_fields:
        subfields = f.get('subFields', [])
        col = []
        for sf in subfields:
            text = sf.get('inferText', '')
            y = sf['boundingPoly']['vertices'][0]['y']
            col.append((y, text))
        col.sort(key=lambda x: x[0])
        cols.append(col)

    # 2. 기준 컬럼(y) 리스트
    ref_col = cols[reference_col_index]
    ref_y_list = [y for y, _ in ref_col]

    # 3. 기준 y에 맞춰 각 컬럼 값 매칭, 없으면 빈 문자열
    rows = []
    for ref_y in ref_y_list:
        row = []
        for col in cols:
            cell = ''
            for y, text in col:
                if abs(y - ref_y) <= y_threshold:
                    cell = text
                    break
            row.append(cell)
        rows.append(row)
    return rows

# 거래내역 테이블 추출
def extract_table_details():
    template_name = template_format.get('template_name')
    fields = template_format.get('fields', [])

    if not fields:
        return [], []

    if template_name == 'bank_shinhan_1':
        # '기본정보' 제외
        target_fields = fields[1:]
    else:
        target_fields = fields

    # 헤더 추출
    thead = [f.get('name') for f in target_fields]
    # row값 추출(y좌표에 따라 셀 구분)
    # tbody = [
    #     (t.split('\n', 1)[1] if '\n' in t else '')  # 첫 번째 \n 이후 부분
    #     for t in (f.get('inferText', '') for f in target_fields)
    # ]
    rows_vertical = rows_by_reference_y_with_text(target_fields, reference_col_index=0, y_threshold=5)

    return thead, rows_vertical

# TODO 숫자 앞 문자 원화 기호로 바꾸기
# TODO 입금, 출금, 잔액 컬럼에는 숫자와 ,(콤마) 만으로 구성,
#     TODO 원화 기호 삭제
# TODO 기본정보가 있으면 기본정보 아래 한 줄 비우고 작성
# TODO 아웃풋 파일을 여러개로 나누기
# TODO df_combined = pd.concat([df_basic_info, df_table], ignore_index=True)
# TODO 기본정보가 있으면 테이블 위에 한 줄 비우고 삽입

def run_ocr_pipeline(pdf_path: str, output_path: str, task_id: str = None,
    progress_callback=None,):
    """
    주어진 PDF를 OCR 처리하여 Excel 파일로 저장하는 메인 함수
    """

    output_dir = os.path.join(os.path.dirname(pdf_path), "temp_split")
    os.makedirs(output_dir, exist_ok=True)

    pdf_pages = split_pdf_pages(pdf_path, output_dir)
    total_pages = len(pdf_pages)
    print(f"{total_pages}개의 페이지로 분할 완료")

    # 시작 시 약간의 진행률 부여 (예: 3%)
    if progress_callback and task_id:
        progress_callback(task_id, 3)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for idx, page_pdf in enumerate(pdf_pages, start=1):
            ocr_result = call_ocr_api(page_pdf)
            if not ocr_result:
                # OCR 실패해도 진행률은 올라가도록 처리
                if progress_callback and task_id:
                    # 전체의 5~95% 구간에서 페이지 단위로 증가
                    progress = int(3 + (idx / total_pages) * 92)  # 3~95 사이
                    progress_callback(task_id, progress)
                continue

            get_template_info(ocr_result)

            basic_info_text = extract_basic_info()
            parsed_basic_info = parse_basic_info(basic_info_text) if basic_info_text else {}
            df_basic_info = pd.DataFrame([parsed_basic_info])

            headers, rows = extract_table_details()
            rows_tbody = rows[1:] if len(rows) > 1 else rows
            df_table = pd.DataFrame(rows_tbody, columns=headers)

            sheet_name = f"Page_{idx}"
            if not df_basic_info.empty:
                df_basic_info.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                start_row = len(df_basic_info) + 2
                df_table.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
            else:
                df_table.to_excel(writer, sheet_name=sheet_name, index=False)

            # 각 페이지 처리 후 진행률 업데이트
            if progress_callback and task_id:
                progress = int(3 + (idx / total_pages) * 92)  # 3~95 사이
                progress_callback(task_id, progress)

    # 마지막으로 100% 보장
    if progress_callback and task_id:
        progress_callback(task_id, 100)

    print(f"✅ 엑셀 파일 저장 완료: {output_path}")
    return output_path

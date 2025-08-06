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
SECRET_KEY = os.getenv("SECRET_KEY")

# PDF -> 단일 페이지 PDF로 분할
def split_pdf_pages(pdf_path, output_dir):
    """
    PDF 파일을 페이지별로 분할하여 개별 PDF 파일로 저장합니다.

    Args:
        pdf_path (str): 원본 PDF 파일의 경로.
        output_dir (str): 분할된 PDF 파일을 저장할 디렉토리.

    Returns:
        list: 분할된 각 페이지 PDF 파일의 경로 리스트.
    """
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
    """
    네이버 CLOVA OCR API를 호출하여 PDF 파일에서 텍스트를 추출합니다.

    Args:
        pdf_path (str): OCR을 수행할 PDF 파일의 경로.

    Returns:
        dict: OCR API 응답 JSON. 오류 발생 시 None.
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
            # print(response.json())
            return response.json()
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] OCR API 호출 실패: {e}")
        return None
    except Exception as e:
        print(f"[ERROR] OCR API 처리 중 예상치 못한 오류 발생: {e}")
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
    """
    OCR 필드 데이터에서 거래내역 테이블을 추출합니다.
    헤더는 subFields에서, 거래내역은 주요 inferText 블록에서 처리합니다.

    Args:
        fields (list): OCR API에서 반환된 모든 필드(텍스트 및 바운딩 박스) 리스트.
        y_threshold (int): (이 함수에서는 직접 사용되지 않지만, 기존 시그니처 유지를 위해 남겨둠)

    Returns:
        tuple: (헤더 리스트, 거래내역 행 리스트).
               각 거래내역 행은 셀 값의 리스트입니다.
    Raises:
        ValueError: 헤더 정보 또는 전체 테이블 텍스트 블록을 찾을 수 없는 경우.
    """
    header_fields_from_subfields = []
    main_table_text_block_content = ""

    # fields 리스트에서 헤더 subFields와 메인 테이블 텍스트 블록을 식별합니다.
    for field_item in fields:
        if 'subFields' in field_item and field_item['subFields']:
            header_fields_from_subfields.extend([
                {'text': sf['inferText'].strip(),
                 'x': min(v.get('x', 0) for v in sf.get('boundingPoly', {}).get('vertices', [])),
                 'y': sum(v.get('y', 0) for v in sf.get('boundingPoly', {}).get('vertices', [])) / 4}
                for sf in field_item['subFields'] if sf.get('inferText') and sf.get('boundingPoly')
            ])
        # 여러 줄을 포함하고 길이가 긴 필드는 전체 테이블 텍스트 블록으로 간주합니다.
        # OCR 응답 샘플에 따르면 'Field 01'이 이 역할을 합니다.
        if '\n' in field_item.get('inferText', '') and len(field_item.get('inferText', '').split('\n')) > 3:
            main_table_text_block_content = field_item['inferText']

    if not header_fields_from_subfields or not main_table_text_block_content:
        raise ValueError("OCR 응답에서 헤더 정보 또는 전체 테이블 텍스트 블록을 찾을 수 없습니다.")

    # 1. 필드 정리 (텍스트 + 중심 좌표 추출) - 이 단계는 이제 header_fields_from_subfields에 직접 적용됩니다.
    # header_fields_from_subfields는 이미 필요한 'text', 'x', 'y' 정보를 가지고 있습니다.

    # 2. y 기준 정렬 및 행 그룹화 - 이 단계는 헤더 식별에만 사용됩니다.
    # (기존 로직 유지)
    items = sorted(header_fields_from_subfields, key=lambda x: x['y'])
    rows_for_header_detection, row, prev_y = [], [], None
    for item in items:
        if prev_y is None or abs(item['y'] - prev_y) <= y_threshold:
            row.append(item)
        else:
            rows_for_header_detection.append(row)
            row = [item]
        prev_y = item['y']
    if row:
        rows_for_header_detection.append(row)

        # 3. 헤더 찾기 (기존 로직 유지)
    for row in rows_for_header_detection:
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

    # 5. x 위치 기준 경계 계산 (기존 로직 유지)
    column_boundaries = []
    if not merged_x: # 헤더가 없는 경우 처리
        return [], []

    for i in range(len(merged_x)):
        left = (merged_x[i - 1] + merged_x[i]) / 2 if i > 0 else -float('inf')
        right = (merged_x[i] + merged_x[i + 1]) / 2 if i < len(merged_x) - 1 else float('inf')
        column_boundaries.append((left, right))

    # 6. 거래내역 행 정렬
    aligned_rows = []
    lines = main_table_text_block_content.split('\n')

    # 실제 데이터가 시작하는 행 찾기 (날짜 패턴으로 식별)
    data_start_index = -1
    for i, line in enumerate(lines):
        # YYYY-MM-DD 또는 YY-MM-DD 또는 YYYY.MM.DD 또는 YY.MM.DD 패턴
        if re.match(r'(\d{4}[-.]\d{2}[-.]\d{2}|\d{2}[-.]\d{2}[-.]\d{2})', line.strip()):
            data_start_index = i
            break

    if data_start_index == -1:
        print("[WARNING] 거래내역 데이터 시작점을 찾을 수 없습니다. 전체 텍스트 블록을 확인하세요.")
        return merged_header, [] # 데이터 행을 찾지 못하면 빈 리스트 반환

    for line_str in lines[data_start_index:]:
        line_str = line_str.strip()
        if not line_str: # 빈 줄 건너뛰기
            continue

        aligned_row = [''] * len(merged_header)

        parts = re.split(r'\s{2,}', line_str)

        # 분리된 부분을 헤더 열에 순차적으로 할당
        for p_idx, part in enumerate(parts):
            if p_idx < len(aligned_row):
                # 기존 내용이 있으면 공백으로 구분하여 추가
                if aligned_row[p_idx]:
                    aligned_row[p_idx] += " " + part
                else:
                    aligned_row[p_idx] = part
            else:
                # 분리된 부분이 헤더 열보다 많으면 마지막 열에 추가 (병합)
                aligned_row[-1] += (" " + part) if aligned_row[-1] else part

        aligned_rows.append(aligned_row)

    return merged_header, aligned_rows

# 숫자 앞 문자 원화 기호로 바꾸기
def replace_W_before_number(text):
    """
    텍스트 내에서 숫자 앞에 있는 'W' 또는 '\' 문자를 '₩' (원화 기호)로 대체합니다.

    Args:
        text (str): 처리할 텍스트.

    Returns:
        str: 원화 기호로 대체된 텍스트.
    """
    return re.sub(r'[W\\](?=\d)', '₩', text)

# Excel 저장
def save_to_excel(wb, sheet_name, basic_info, header, rows):
    """
    추출된 기본 정보와 거래내역을 Excel 워크북에 시트로 저장합니다.

    Args:
        wb (openpyxl.Workbook): 저장할 Excel 워크북 객체.
        sheet_name (str): 생성할 시트의 이름.
        basic_info (dict): 기본 정보 딕셔너리.
        header (list): 테이블 헤더 리스트.
        rows (list): 테이블 데이터 행 리스트.
    """
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
    """
    PDF 파일을 처리하고 OCR을 수행하여 Excel 파일로 저장하는 메인 함수입니다.

    Args:
        input_pdf_path (str): 처리할 입력 PDF 파일의 경로.
        update_progress_callback (function, optional): 진행률을 업데이트하기 위한 콜백 함수.
                                                      Defaults to None.

    Returns:
        openpyxl.Workbook: 처리된 데이터가 포함된 Excel 워크북 객체.
    """
    temp_dir = os.path.join('temp_split', os.path.splitext(os.path.basename(input_pdf_path))[0])

    pdf_pages = split_pdf_pages(input_pdf_path, temp_dir)
    total_pages = len(pdf_pages)

    wb = Workbook()
    # 기본으로 생성된 시트 제거
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    for idx, page_pdf in enumerate(pdf_pages):
        print(f"📄 처리 중: 페이지 {idx + 1}/{total_pages}")

        result = call_ocr_api(page_pdf)
        if not result or 'images' not in result:
            print(f"[⚠️] 페이지 {idx + 1} OCR 실패 또는 응답에 'images' 키 없음.")
            if update_progress_callback:
                update_progress_callback((idx + 1) / total_pages * 100)
            continue

        # 모든 필드를 하나의 리스트로 모읍니다.
        all_fields_from_ocr = []
        full_page_text_for_basic_info = ""

        for image_data in result['images']:
            if 'fields' in image_data:
                all_fields_from_ocr.extend(image_data['fields'])
                for field_item in image_data['fields']:
                    full_page_text_for_basic_info += field_item.get('inferText', '') + " "

            if 'title' in image_data and 'subFields' in image_data['title']:
                # title 필드 자체도 all_fields_from_ocr에 포함될 수 있습니다.
                # 여기서는 subFields만 추출하여 header_fields_from_subfields로 넘겨줄 필요가 없으므로,
                # 단순히 full_page_text_for_basic_info에 추가합니다.
                full_page_text_for_basic_info += image_data['title'].get('inferText', '') + " "


        # 페이지 처리 진행률 업데이트
        if update_progress_callback:
            progress = (idx + 1) / total_pages * 100
            update_progress_callback(int(progress))



        # 기본 정보 추출
        basic_info = extract_basic_info(full_page_text_for_basic_info)
        print(f'페이지 {idx+1}에서 추출된 텍스트 (일부): {full_page_text_for_basic_info[:200]}...') # 디버깅을 위해 처음 200자 인쇄

        try:
            # extract_transaction_table_by_position 함수에 모든 필드를 전달합니다.
            header, rows = extract_transaction_table_by_position(all_fields_from_ocr)
            save_to_excel(wb, f"페이지 {idx + 1}", basic_info, header, rows)
            print(f"✅ 페이지 {idx + 1} 거래 내역 추출 및 Excel 저장 완료.")
        except ValueError as ve:
            print(f"[❌] 페이지 {idx + 1} 거래 내역 처리 오류: {ve}")
        except Exception as e:
            print(f"[❌] 페이지 {idx + 1} 처리 중 예상치 못한 오류 발생: {e}")

    return wb

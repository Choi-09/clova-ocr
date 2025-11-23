import re
from typing import List, Tuple

import pandas as pd

from .ocr_common import run_ocr_pipeline_core


def get_template_info(result: dict) -> Tuple[int, list]:
    """
    OCR 결과에서 템플릿 이름과 필드 목록 추출
    """
    template_id = 0
    fields = []
    try:
        images = result.get("images", None)
        if images:
            image = images[0]
            matched_template = image.get("matchedTemplate", {})
            template_id = matched_template.get("id", "")
            fields = image.get("fields", [])
    except Exception as e:
        print(f"[ERROR] 템플릿 정보 추출 실패: {e}")

    return template_id, fields


def extract_basic_info(template_id: int, fields: list):
    """
    기본 정보 필드 추출
    - 신한 1페이지 템플릿: template_id == 39249'
    - 첫 번째 필드 이름이 '기본정보' 인 경우만 처리
    """
    if template_id == 39249 and fields:
        first_field = fields[0]
        if first_field.get("name") == "기본정보":
            # TODO first_field.get('subFields') 에서 'inferText' 내용 추출,
            #  'boundingPoly'의 좌표 추출해서 dataFrame으로 정렬=> df리턴
            basic_info = first_field.get("inferText")
            return basic_info

    # template_id가 다르거나 기본정보가 없으면 None
    return None


def parse_basic_info(basic_info_text: str):
    """
    OCR 기본정보 텍스트를 각 항목별로 분리
    """
    if not basic_info_text:
        return {}

    # 한 줄로 합치기
    text = basic_info_text.replace("\n", " ")

    # 추출할 항목 리스트
    keys = ["계좌번호", "조회기간", "예금주 성명", "상품명"]

    result = {}
    for i, key in enumerate(keys):
        if i < len(keys) - 1:
            next_key = keys[i + 1]
            pattern = rf"{key}\s+(.*?)(?=\s{next_key})"
        else:
            pattern = rf"{key}\s+(.*)"

        match = re.search(pattern, text)
        if match:
            value = match.group(1).strip()
            # 조회기간 처리: 두 날짜 사이에 ~ 삽입
            if key == "조회기간":
                dates = value.split()
                if len(dates) == 2:
                    value = f"{dates[0]} ~ {dates[1]}"
            result[key] = value
    return result


def rows_by_reference_y_with_text(
    target_fields: list,
    reference_col_index: int = 0,
    y_threshold: int = 5,
) -> List[List[str]]:
    """
    거래일자 컬럼 기준으로 y 좌표를 맞추고, 각 셀은 첫 번째 \n 이후 텍스트를 사용
    """
    # 1. 각 컬럼별 (y, text) 리스트 생성
    cols = []
    for f in target_fields:
        subfields = f.get("subFields", [])
        col = []
        for sf in subfields:
            text = sf.get("inferText", "")
            y = sf["boundingPoly"]["vertices"][0]["y"]
            col.append((y, text))
        col.sort(key=lambda x: x[0])
        cols.append(col)

    # 2. 기준 컬럼(y) 리스트
    ref_col = cols[reference_col_index]
    ref_y_list = [y for y, _ in ref_col]

    # 3. 기준 y에 맞춰 각 컬럼 값 매칭, 없으면 빈 문자열
    # TODO 숫자 앞 문자 원화 기호로 바꾸기
    # TODO 입금, 출금, 잔액 컬럼에는 숫자와 ,(콤마) 만으로 구성,
    # TODO 원화 기호 삭제

    rows = []
    for ref_y in ref_y_list:
        row = []
        for col in cols:
            cell = ""
            for y, text in col:
                if abs(y - ref_y) <= y_threshold:
                    cell = text
                    break
            row.append(cell)
        rows.append(row)
    return rows


def extract_table_details(template_id:int, fields: list):
    """
    거래내역 테이블 헤더/바디 추출
    """
    if not fields:
        return [], []

    if template_id == 39249:
        # '기본정보' 제외
        target_fields = fields[1:]
    else:
        target_fields = fields

    # 헤더 추출
    thead = [f.get("name") for f in target_fields]

    # y좌표 기준으로 행 구성
    rows_vertical = rows_by_reference_y_with_text(
        target_fields, reference_col_index=0, y_threshold=5
    )

    return thead, rows_vertical


def parse_shinhan_page(ocr_result: dict, page_index: int):
    """
    run_ocr_pipeline_core 에 넘길 페이지 파서
    - ocr_result 를 받아 기본정보/거래내역 DataFrame 두 개를 반환
    """
    template_id, fields = get_template_info(ocr_result)

    # 기본정보
    basic_info_text = extract_basic_info(template_id, fields)
    parsed_basic_info = parse_basic_info(basic_info_text) if basic_info_text else {}

    if parsed_basic_info:
        df_basic_info = pd.DataFrame([parsed_basic_info])
    else:
        df_basic_info = pd.DataFrame()

    # 테이블
    headers, rows = extract_table_details(template_id, fields)
    rows_tbody = rows[1:] if len(rows) > 1 else rows

    if headers:
        df_table = pd.DataFrame(rows_tbody, columns=headers)
    else:
        df_table = pd.DataFrame(rows_tbody)

    return df_basic_info, df_table


def run_ocr_pipeline(
    pdf_path: str,
    output_path: str,
    task_id: str = None,
    progress_callback=None,
):
    """
    신한은행 전용 OCR 파이프라인
    - 공통 모듈의 run_ocr_pipeline_core 를 호출
    - 페이지 파서로 parse_shinhan_page 사용
    """
    return run_ocr_pipeline_core(
        pdf_path=pdf_path,
        output_path=output_path,
        page_parser=parse_shinhan_page,
        task_id=task_id,
        progress_callback=progress_callback,
    )

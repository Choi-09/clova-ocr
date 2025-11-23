import json
import os
import time
import uuid
from typing import Callable, Tuple

import pandas as pd
import requests
from PyPDF2 import PdfReader, PdfWriter
from dotenv import load_dotenv

load_dotenv()

API_URL = os.getenv("API_URL")
SECRET_KEY = os.getenv("SECRET_KEY")


def split_pdf_pages(pdf_path: str, output_dir: str):
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


def call_ocr_api(pdf_path: str):
    """
    API를 호출하여 PDF 파일에서 텍스트 추출
    """
    try:
        with open(pdf_path, "rb") as file:
            payload = {
                "message": json.dumps(
                    {
                        "images": [{"format": "pdf", "name": "demo"}],
                        "requestId": str(uuid.uuid4()),
                        "version": "V2",
                        "timestamp": int(time.time() * 1000),
                    }
                ).encode("UTF-8")
            }
            files = [("file", file)]
            headers = {"X-OCR-SECRET": SECRET_KEY}
            response = requests.post(API_URL, headers=headers, data=payload, files=files)
            response.raise_for_status()
            return response.json()
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] OCR API 호출 실패: {e}")
        return None
    except Exception as e:
        print(f"[ERROR] OCR API 처리 중 예상치 못한 오류 발생: {e}")
        return None


# 페이지별 파서 타입: (ocr_result, page_index) -> (df_basic_info, df_table)
PageParser = Callable[[dict, int], Tuple[pd.DataFrame, pd.DataFrame]]


def run_ocr_pipeline_core(
    pdf_path: str,
    output_path: str,
    page_parser: PageParser,
    task_id: str = None,
    progress_callback=None,
):
    """
    공통 OCR 파이프라인:
    - PDF 페이지 분할
    - 각 페이지별 OCR API 호출
    - page_parser 로 DataFrame 생성
    - Excel Writer 에 페이지별 시트 저장

    page_parser 는 (ocr_result, page_index) -> (df_basic_info, df_table) 형태의 콜백
    """
    # TODO 기본정보가 있으면 기본정보 아래 한 줄 비우고 작성
    # TODO 아웃풋 파일을 여러개로 나누기
    # TODO df_combined = pd.concat([df_basic_info, df_table], ignore_index=True)
    # TODO 기본정보가 있으면 테이블 위에 한 줄 비우고 삽입
    # 페이지 분할
    output_dir = os.path.join(os.path.dirname(pdf_path), "temp_split")
    os.makedirs(output_dir, exist_ok=True)

    pdf_pages = split_pdf_pages(pdf_path, output_dir)
    total_pages = len(pdf_pages)
    print(f"{total_pages}개의 페이지로 분할 완료")

    # 시작 시 약간의 진행률 부여 (예: 3%)
    if progress_callback and task_id:
        progress_callback(task_id, 10)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for idx, page_pdf in enumerate(pdf_pages, start=1):
            ocr_result = call_ocr_api(page_pdf)
            if not ocr_result:
                # OCR 실패해도 진행률은 올라가도록 처리
                if progress_callback and task_id:
                    progress = int(3 + (idx / total_pages) * 92)  # 3~95 사이
                    progress_callback(task_id, progress)
                continue

            # 은행별/양식별 페이지 파서 호출
            df_basic_info, df_table = page_parser(ocr_result, idx)

            # None 방어
            if df_basic_info is None:
                df_basic_info = pd.DataFrame()
            if df_table is None:
                df_table = pd.DataFrame()

            sheet_name = f"Page_{idx}"
            if not df_basic_info.empty:
                # 기본정보 먼저
                df_basic_info.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                # 한 줄 비우고 테이블 작성
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

# Clova-OCR

Clova OCR 기반 은행 거래 내역 PDF 자동 분석 및 엑셀 변환 웹 애플리케이션입니다.

---

## 주요 기능

- PDF 은행 거래 내역 파일 업로드
- PDF 각 페이지별 OCR 처리 및 거래 내역 추출
- 결과를 엑셀 파일로 생성 및 다운로드 제공
- 업로드 및 처리 진행률 표시 (비동기 프로그래스바)
- Flask 기반 웹 서버와 간단한 HTML 프론트엔드 제공

---

## 설치 및 실행 방법

### 1. 필수 조건
- Python 3.8 이상
- pip
- Redis (선택 사항, 진행률 표시용)

### 2. 패키지 설치
```bash
pip install -r requirements.txt
```

### 3. 프로젝트 실행
```bash
python app.py
```
- 웹 서버가 http://127.0.0.1:5000 에서 실행됩니다.

---

## 사용법
#### 1. 웹페이지에서 PDF 파일을 선택해 업로드합니다.
#### 2. 서버에서 OCR 및 데이터 추출이 진행됩니다.
#### 3. 진행률이 실시간 표시됩니다.
#### 4. 처리 완료 후 결과 엑셀 파일을 다운로드할 수 있습니다.

---

## 주요 파일
- app.py: Flask 웹 서버 및 라우팅 구현
- utils/bank_shinhan_process.py: PDF 분할, OCR API 호출, 엑셀 저장 등 핵심 처리 함수 포함
- templates/: HTML 템플릿 파일 (index.html, result.html)
- static/: 정적 파일 (CSS, JS 등)

---

## 진행률 표시
- process_pdf_and_ocr 함수 내에서 페이지별로 진행률 콜백 호출

---


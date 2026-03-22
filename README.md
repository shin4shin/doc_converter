# 문서 변환기

![version](https://img.shields.io/badge/version-v1.1.0-blue) ![platform](https://img.shields.io/badge/platform-Windows-lightgrey)

HWP, HWPX, DOCX, PDF 파일을 서로 변환할 수 있는 Windows용 데스크톱 애플리케이션입니다.

---

## 주요 기능

- HWP / HWPX → PDF, DOCX, TXT 변환
- DOCX → PDF, TXT, HWP, HWPX 변환
- PDF → TXT, DOCX, HWP, HWPX 변환
- 직관적인 GUI 제공 (tkinter 기반)
- 변환 완료 후 저장 폴더 자동 열기

---

## 지원 변환 목록

| 입력 형식 | 출력 가능 형식 |
|-----------|----------------|
| HWP       | PDF, DOCX, TXT |
| HWPX      | PDF, DOCX, TXT |
| DOCX      | PDF, TXT, HWP, HWPX |
| PDF       | TXT, DOCX, HWP, HWPX |

> HWP / HWPX로 저장 시 실제로는 HWPX 포맷으로 생성됩니다.  
> HWPX는 한글 2019 이상에서 열 수 있습니다.

---

## 설치 및 실행

### 이용자

1. [LibreOffice](https://www.libreoffice.org/download/download/) 설치
2. `문서변환기.exe` 더블클릭

Python 설치 불필요. LibreOffice만 있으면 됩니다.

### 개발자 (소스에서 실행)

**요구 사항**

- Python 3.8 이상
- LibreOffice

**패키지 설치**

```bash
pip install python-docx pypdf pdfplumber reportlab pyhwp lxml
```

**실행**

```bash
python doc_converter_all.py
```

**exe 빌드**

`빌드.bat` 을 더블클릭하면 자동으로 패키지 설치 및 빌드가 진행됩니다.  
완료 후 `dist/문서변환기.exe` 가 생성됩니다.

---

## 파일 구성

```
문서변환기/
├── doc_converter_all.py   # 변환 엔진 + GUI 통합 소스
├── 빌드.bat               # PyInstaller exe 빌드 스크립트
└── 사용안내.txt           # 이용자용 설치 안내
```

---

## 기술 스택

| 역할 | 라이브러리 |
|------|-----------|
| GUI | tkinter |
| HWP 읽기 | pyhwp |
| HWPX 읽기/쓰기 | lxml + zipfile (직접 파싱) |
| DOCX 처리 | python-docx |
| PDF 읽기 | pdfplumber, pypdf |
| PDF 생성 | reportlab |
| HWP/DOCX/PDF 변환 | LibreOffice (headless) |
| exe 빌드 | PyInstaller |

---

## 변환 방식

- **HWP → PDF / DOCX** : LibreOffice headless 변환
- **HWP → TXT** : pyhwp (hwp5txt) 직접 추출
- **HWPX → TXT** : ZIP 압축 해제 후 XML 직접 파싱
- **HWPX → PDF / DOCX** : LibreOffice 변환 시도 → 실패 시 텍스트 기반 폴백
- **DOCX → PDF** : LibreOffice headless 변환
- **DOCX → HWPX** : python-docx로 단락·표 추출 후 HWPX XML 직접 생성
- **PDF → TXT** : pdfplumber 추출
- **PDF → DOCX** : LibreOffice headless 변환
- **PDF → HWPX** : PDF → DOCX → HWPX 순차 변환

---

## 주의 사항

- DOCX / PDF → HWP 변환은 HWPX 포맷으로 저장됩니다. HWP 바이너리 포맷은 직접 생성이 불가능합니다.
- 스캔된 PDF(이미지 기반)는 텍스트 추출이 되지 않습니다.
- 이미지, 복잡한 서식은 변환 과정에서 일부 손실될 수 있습니다.
- LibreOffice가 설치되어 있지 않으면 HWP, HWPX, DOCX ↔ PDF 변환이 동작하지 않습니다.

## 변경 이력

### v1.1.0 (2026-03-22)
- DOCX → HWP / HWPX 변환 추가
- PDF → HWP / HWPX 변환 추가
- GUI 버튼 활성화 버그 수정 (docx, pdf 선택 시 hwp/hwpx 버튼 비활성화 문제)
- Windows LibreOffice 경로 자동 탐색 개선 (glob 패턴으로 버전 무관 탐색)
- 변환 엔진과 GUI를 단일 파일(doc_converter_all.py)로 통합
- import 오류 방지를 위한 의존성 검증 로직 추가

### v1.0.0
- 최초 릴리스
- HWP / HWPX → PDF, DOCX, TXT 변환 지원
- DOCX → PDF, TXT 변환 지원
- PDF → TXT, DOCX 변환 지원
- tkinter 기반 GUI 제공

---

## 라이선스

MIT License

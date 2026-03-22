@echo off
chcp 65001 > nul
title 문서 변환기

:: ── Python 찾기 ──────────────────────────────
set PYTHON=
for %%P in (python.exe) do set PYTHON=%%~$PATH:P

if not defined PYTHON (
    :: 일반적인 설치 경로 직접 탐색
    for %%V in (313 312 311 310 39 38) do (
        if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
            set PYTHON=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe
            goto :found_python
        )
    )
    echo [오류] Python 을 찾을 수 없습니다.
    echo Python 을 설치하세요: https://www.python.org/downloads/
    pause
    exit /b 1
)

:found_python
:: ── doc_converter_all.py 찾기 ────────────────
:: 1순위: 이 bat 파일과 같은 폴더
set SCRIPT=%~dp0doc_converter_all.py

if not exist "%SCRIPT%" (
    :: 2순위: 바탕화면
    set SCRIPT=%USERPROFILE%\Desktop\doc_converter_all.py
)

if not exist "%SCRIPT%" (
    :: 3순위: OneDrive 바탕화면
    set SCRIPT=%USERPROFILE%\OneDrive\바탕 화면\doc_converter_all.py
)

if not exist "%SCRIPT%" (
    echo [오류] doc_converter_all.py 를 찾을 수 없습니다.
    echo 이 bat 파일과 같은 폴더에 doc_converter_all.py 를 넣어주세요.
    pause
    exit /b 1
)

:: ── 필수 패키지 자동 설치 ────────────────────
echo 패키지 확인 중...
"%PYTHON%" -c "import docx, pypdf, pdfplumber, reportlab, hwp5, lxml" 2>nul
if errorlevel 1 (
    echo 필요한 패키지를 설치합니다. 잠시 기다려주세요...
    "%PYTHON%" -m pip install python-docx pypdf pdfplumber reportlab pyhwp lxml --quiet
    if errorlevel 1 (
        echo [경고] 일부 패키지 설치에 실패했습니다. 변환 시 오류가 발생할 수 있습니다.
    ) else (
        echo 패키지 설치 완료.
    )
)

:: ── 실행 ─────────────────────────────────────
echo 문서 변환기를 실행합니다...
"%PYTHON%" "%SCRIPT%"
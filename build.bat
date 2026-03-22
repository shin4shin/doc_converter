@echo off
chcp 65001 > nul
title 문서변환기 빌드

echo ================================================
echo   문서변환기 exe 빌드 스크립트
echo ================================================
echo.

:: ── Python 찾기 ──────────────────────────────────
set PYTHON=
for %%P in (python.exe) do set PYTHON=%%~$PATH:P

if not defined PYTHON (
    for %%V in (313 312 311 310 39 38) do (
        if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
            set PYTHON=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe
            goto :found_python
        )
    )
    echo [오류] Python 을 찾을 수 없습니다.
    pause & exit /b 1
)

:found_python
echo [1/4] Python 확인: %PYTHON%

:: ── 필수 패키지 설치 ─────────────────────────────
echo [2/4] 필수 패키지 설치 중...
"%PYTHON%" -m pip install ^
    pyinstaller ^
    python-docx ^
    pypdf ^
    pdfplumber ^
    reportlab ^
    pyhwp ^
    lxml ^
    --quiet

if errorlevel 1 (
    echo [오류] 패키지 설치 실패
    pause & exit /b 1
)
echo        완료.

:: ── doc_converter_all.py 확인 ───────────────────
set SCRIPT=%~dp0doc_converter_all.py
if not exist "%SCRIPT%" (
    echo [오류] doc_converter_all.py 를 이 파일과 같은 폴더에 넣어주세요.
    pause & exit /b 1
)
echo [3/4] 소스 파일 확인: %SCRIPT%

:: ── PyInstaller 빌드 ─────────────────────────────
echo [4/4] exe 빌드 중... (1~3분 소요)
echo.

"%PYTHON%" -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "문서변환기" ^
    --distpath "%~dp0dist" ^
    --workpath "%~dp0build" ^
    --specpath "%~dp0build" ^
    --hidden-import docx ^
    --hidden-import pypdf ^
    --hidden-import pdfplumber ^
    --hidden-import reportlab ^
    --hidden-import reportlab.lib.pagesizes ^
    --hidden-import reportlab.platypus ^
    --hidden-import reportlab.lib.styles ^
    --hidden-import reportlab.pdfbase ^
    --hidden-import reportlab.pdfbase.ttfonts ^
    --hidden-import hwp5 ^
    --hidden-import lxml ^
    --hidden-import lxml.etree ^
    --hidden-import zipfile ^
    --hidden-import threading ^
    --hidden-import tkinter ^
    --hidden-import tkinter.ttk ^
    --hidden-import tkinter.filedialog ^
    --hidden-import tkinter.messagebox ^
    --collect-all pdfplumber ^
    --collect-all hwp5 ^
    "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패. 위 오류 메시지를 확인해주세요.
    pause & exit /b 1
)

echo.
echo ================================================
echo   빌드 완료!
echo   파일 위치: %~dp0dist\문서변환기.exe
echo ================================================
echo.
echo 이용자에게 배포할 파일:
echo   dist\문서변환기.exe
echo.
echo ※ 이용자는 LibreOffice 만 설치하면 됩니다.
echo   https://www.libreoffice.org/download/download/
echo.

:: dist 폴더 열기
explorer "%~dp0dist"
pause
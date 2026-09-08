@echo off
title ARGOS.UZ - Web Vakansiyalar Tahlilchisi va Word Viewer
cd /d "%~dp0"

echo ========================================================================
echo   ARGOS.UZ - WEB VAKANSIYALAR TAHLILCHISI VA WORD VIEWER
echo ========================================================================
echo.

:: 1. Python dasturini aniqlash
set "PY_CMD=python"
where python >nul 2>nul
if errorlevel 1 (
    where py >nul 2>nul
    if not errorlevel 1 (
        set "PY_CMD=py"
    ) else (
        echo [XATO] Kompyuteringizda Python topilmadi!
        echo Iltimos, Python 3.10+ o'rnatilganligiga va PATH ga qo'shilganligiga ishonch hosil qiling.
        echo https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

echo [*] Python aniqlandi: %PY_CMD%

:: 2. Kerakli kutubxonalarni tekshirish
echo [*] Kutubxonalar tekshirilmoqda...
%PY_CMD% -c "import flask, mammoth, docx, requests, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [*] Kutubxonalar yetishmayapti, avtomatik o'rnatilmoqda...
    %PY_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [OGOHLANTIRISH] pip install xatolik berdi, qayta urinib ko'rilmoqda...
        %PY_CMD% -m pip install flask mammoth python-docx requests openpyxl
    )
)

:: 3. Serverni ishga tushirish
echo.
echo ========================================================================
echo [*] Web server ishga tushirilmoqda...
echo [*] Manzil: http://127.0.0.1:5000
echo [*] Brauzeringiz avtomatik tarzda ochiladi.
echo [*] Serverni to'xtatish uchun: ushbu oynani yoping yoki Ctrl + C bosing.
echo ========================================================================
echo.

%PY_CMD% app.py

if errorlevel 1 (
    echo.
    echo [XATO] Web server kutilmaganda to'xtadi!
    pause
)


@echo off
chcp 65001 >nul
cd /d "%~dp0"

title HLR Checker

echo ============================================================
echo   HLR Checker - запуск
echo ============================================================
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ОШИБКА] Python не найден.
    echo Установите Python 3 с сайта python.org
    echo и отметьте галочку "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

echo [1/2] Устанавливаю зависимости...
python -m pip install -r requirements.txt --quiet

echo [2/2] Запускаю программу...
echo.
python main.py

pause


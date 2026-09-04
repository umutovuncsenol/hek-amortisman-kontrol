@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo HEK Amortisman Kontrol uygulamasi hazirlaniyor...
echo.

where py >nul 2>&1
if %errorlevel%==0 (
    set "APP_PYTHON=py -3"
) else (
    where python >nul 2>&1
    if errorlevel 1 (
        echo HATA: Python bulunamadi.
        pause
        exit /b 1
    )
    set "APP_PYTHON=python"
)

%APP_PYTHON% -m venv .build-venv
if errorlevel 1 goto :error
call .build-venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install -r requirements.txt pyinstaller==6.22.2
if errorlevel 1 goto :error

python -m PyInstaller --noconfirm --clean --onefile --windowed --noupx --collect-all openpyxl --add-data "assets\renault-logo-header.png;assets" --add-data "assets\renault-logo.png;assets" --add-data "assets\HEK_Amortisman_Ornek_Cikti.xlsx;assets" --icon "assets\renault-logo.ico" --name HEKAmortismanKontrol hek_amortisman_kontrol.py
if errorlevel 1 goto :error

echo.
echo TAMAMLANDI: dist\HEKAmortismanKontrol.exe
copy /Y KULLANIM.txt dist\KULLANIM.txt >nul
powershell -NoProfile -Command "Compress-Archive -Force -Path 'dist\HEKAmortismanKontrol.exe','dist\KULLANIM.txt' -DestinationPath 'dist\HEKAmortismanKontrol-Windows.zip'"
start "" "%CD%\dist"
pause
exit /b 0

:error
echo.
echo HATA: Uygulama olusturulamadi.
pause
exit /b 1

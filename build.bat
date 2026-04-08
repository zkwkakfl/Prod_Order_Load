@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo [1/4] Python 확인...
python --version >nul 2>&1
if errorlevel 1 (
  echo Python이 PATH에 없습니다. python.org 에서 3.10+ 설치 후 다시 실행하세요.
  exit /b 1
)

echo [2/4] 가상환경 ^(.venv^) 준비...
if not exist ".venv\Scripts\python.exe" (
  python -m venv .venv
  if errorlevel 1 exit /b 1
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 exit /b 1

python -m pip install -q -U pip
echo [3/4] 의존성 및 PyInstaller 설치...
pip install -q -r requirements.txt
if errorlevel 1 exit /b 1
pip install -q pyinstaller
if errorlevel 1 exit /b 1

echo [4/4] 단일 exe 빌드 ^(GUI, 콘솔 없음^)...
pyinstaller --noconfirm --clean --windowed --onefile ^
  --name "Prod_Order_Load" ^
  --add-data "VERSION;." ^
  app.py
if errorlevel 1 (
  echo 빌드 실패.
  exit /b 1
)

echo.
echo 완료: dist\Prod_Order_Load.exe
echo 소스 경로 설정^(source_paths.json^)과 결과 파일은 exe와 같은 폴더에 둡니다.
endlocal
exit /b 0

@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  SPC 자료 메일 수집 (최근 90일, 수집함에 저장)
echo  원본 SPC 폴더와 메일은 건드리지 않습니다.
echo  Outlook 이 켜져 있어야 합니다. 몇 분 걸립니다.
echo ============================================
where uv >nul 2>nul
if %errorlevel%==0 (
  uv run --with pywin32 --with xlrd --with openpyxl --with pyzipper python ingest_spc.py --days 90 > _실행결과.txt 2>&1
) else (
  python -m pip install -q pywin32 xlrd openpyxl pyzipper
  python ingest_spc.py --days 90 > _실행결과.txt 2>&1
)
echo.
echo 끝났습니다. 이 창은 닫아도 됩니다.
pause

@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist spc_mail_config.json (
  echo spc_mail_config.json 이 없습니다. Claude 에게 알려주세요.
  pause
  exit /b 1
)
echo ============================================
echo  SPC 자료 메일 분류 시험 (저장 안 함, 최근 90일)
echo  Outlook 이 켜져 있어야 합니다. 몇 분 걸립니다.
echo ============================================
where uv >nul 2>nul
if %errorlevel%==0 (
  uv run --with pywin32 --with xlrd --with openpyxl python ingest_spc.py --dry-run --days 90 > _시험결과.txt 2>&1
) else (
  python -m pip install -q pywin32 xlrd openpyxl
  python ingest_spc.py --dry-run --days 90 > _시험결과.txt 2>&1
)
echo.
echo 끝났습니다. 결과는 Claude 가 직접 확인합니다. 이 창은 닫아도 됩니다.
pause

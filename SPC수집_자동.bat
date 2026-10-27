@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem 작업 스케줄러용: auto_run.py(수집, 자동 결산, 오후 과거 메일)를 pythonw로 창 없이 띄우고 바로 끝낸다
rem pythonw가 없으면 이 창에서 실행한다(창을 닫으면 끊기니 닫지 않기)
set LOG=..\수집함\_자동실행로그.txt
where pythonw >nul 2>nul
if errorlevel 1 goto window
echo %date% %time% 예약 실행 시작 - 창 없이 >> "%LOG%"
start "" pythonw "%~dp0auto_run.py"
exit /b 0
:window
echo %date% %time% 예약 실행 시작 - pythonw가 없어 창에서 실행 >> "%LOG%"
python "%~dp0auto_run.py"
exit /b 0

$ErrorActionPreference = 'Stop'
$bat = Join-Path $env:USERPROFILE 'Desktop\용역\spc 자동화\수집\SPC수집_자동.bat'
if (-not (Test-Path $bat)) { Write-Host "실행 파일을 찾지 못했습니다: $bat"; exit 1 }
$action   = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument ('/c "' + $bat + '"')
$days     = 'Monday','Tuesday','Wednesday','Thursday','Friday'
$trigger1 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At '09:05'
$trigger2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $days -At '14:05'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 1)
Register-ScheduledTask -TaskName 'SPC자료수집' -Action $action -Trigger $trigger1,$trigger2 -Settings $settings `
  -Description 'Outlook 메일의 SPC 자료(자금판·거래내역 등)를 수집함에 자동 분류. 평일 09:05, 14:05.' -Force | Out-Null
Write-Host ''
Write-Host '등록했습니다: 평일 오전 9시 5분, 오후 2시 5분 (Outlook 이 켜진 상태에서 실행)'
Write-Host '기록: 바탕화면\용역\spc 자동화\수집함\_자동실행로그.txt'

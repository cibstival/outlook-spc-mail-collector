Unregister-ScheduledTask -TaskName 'SPC자료수집' -Confirm:$false -ErrorAction SilentlyContinue
Write-Host '매일 자동 수집을 해제했습니다.'

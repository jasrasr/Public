@echo off
cd C:\users\%username%\Desktop\
md Speedy
cd Speedy

:: Add a timestamp header
echo --- %date% %time% --- >> C:\users\%username%\Desktop\Speedy\wifi_speed_log.json

:: Run the speed test and append JSON
speedtest.exe --format=json >> C:\users\%username%\Desktop\Speedy\wifi_speed_log.json

:: Parse the most recent result and append summary
powershell -Command ^
  "$json = Get-Content 'C:\users\%username%\Desktop\Speedy\wifi_speed_log.json' | Select-String '{' -Context 0,100 | Select -Last 1 | ForEach-Object { $_.Line } | ConvertFrom-Json;" ^
  "$down = [math]::Round(($json.download.bandwidth * 8) / 1MB, 2);" ^
  "$up = [math]::Round(($json.upload.bandwidth * 8) / 1MB, 2);" ^
  "$summary = ('Down: {0} Mbps | Up: {1} Mbps' -f $down, $up);" ^
  "Add-Content 'C:\users\%username%\Desktop\Speedy\wifi_speed_log.json' $summary"

:: Add an empty line for readability
echo. >> C:\users\%username%\Desktop\Speedy\wifi_speed_log.json

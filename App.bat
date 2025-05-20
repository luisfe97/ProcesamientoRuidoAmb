@echo off
powershell -windowstyle hidden -command "cd '%~dp0'; Start-Process -NoNewWindow -FilePath '%LOCALAPPDATA%\Microsoft\WindowsApps\python3.13.exe' -ArgumentList 'run_gui.py'"
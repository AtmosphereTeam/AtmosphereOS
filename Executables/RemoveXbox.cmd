@echo off
cd /d "C:\Program Files\WindowsApps\"
for /d %a in (*xbox*) do if exist "%a" rmdir /s /q "%a"
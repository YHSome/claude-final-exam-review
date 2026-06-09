@echo off
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8888 ^| findstr LISTENING') do (
    echo Killing PID %%a on port 8888...
    taskkill /PID %%a /F
)
echo Done.
pause

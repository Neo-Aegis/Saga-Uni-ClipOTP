@echo off
schtasks /delete /tn "SagaOTP_Watcher" /f
if %errorlevel% equ 0 (
    echo OK: Task removed.
) else (
    echo FAILED: Task not found or run as administrator.
)
pause

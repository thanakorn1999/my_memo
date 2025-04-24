@echo off

:: Check for admin rights
net session >nul 2>&1
if %errorLevel% == 0 (
    :: Already running as admin
    cd /d C:\Users\Jame\Desktop\code\me\my_memo
    if %errorLevel% neq 0 (
        echo Failed to change directory. Ensure the path is correct.
        pause
        exit /B
    )
    call env\Scripts\activate.bat
    if %errorLevel% neq 0 (
        echo Failed to activate virtual environment. Ensure the path is correct.
        pause
        exit /B
    )
    python tools.py --t=RPGG
    if %errorLevel% neq 0 (
        echo Failed to run Python script. Ensure Python is installed and the script path is correct.
        pause
        exit /B
    )
) else (
    :: Not running as admin, so re-launch as admin
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %*", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B
)

:: Keep the command prompt open after execution
cmd /k

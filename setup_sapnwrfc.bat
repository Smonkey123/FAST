@echo off

REM Move folder
if exist nwrfc750P_5-70002752 (
    echo Moving nwrfc750P_5-70002752 folder to C:\Temp ...
    robocopy nwrfc750P_5-70002752 C:\Temp\nwrfc750P_5-70002752 /E /MOVE
) else (
    echo Error: nwrfc750P_5-70002752 folder does not exist
    goto :eof
)

REM Set system environment variable
echo Setting system environment variable SAPNWRFC_HOME ...
reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v SAPNWRFC_HOME /t REG_SZ /d "C:\Temp\nwrfc750P_5-70002752\nwrfcsdk" /f

if %errorlevel% neq 0 (
    echo Error: Failed to set environment variable.
) else (
    echo Environment variable set successfully.
)

echo Operation completed.
pause
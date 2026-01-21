@echo off
echo Creating registry entries...

REM Create MSIPC node
reg add "HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\MSIPC" /f

REM Create aip-addin node
reg add "HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\MSIPC\aip-addin" /f

REM Create RMSUser value under aip-addin node
reg add "HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\MSIPC\aip-addin" /v "RMSUser" /t REG_SZ /d "livia-wenjia.qi@cn.abb.com" /f

REM Create UPN value under aip-addin node
reg add "HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\MSIPC\aip-addin" /v "UPN" /t REG_SZ /d "livia-wenjia.qi@cn.abb.com" /f

echo Registry entries created successfully.
pause

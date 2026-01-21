@echo off
REM 检查C:\Temp文件夹是否存在，如果不存在则创建
if not exist "C:\Temp" (
    mkdir "C:\Temp"
)

REM 复制Template文件夹及其内容到C:\Temp
xcopy "Template" "C:\Temp\Template" /E /I /Y

REM 显示复制完成的提示
echo 文件夹复制完成！
pause
@echo off
chcp 65001 >nul
echo ========================================
echo    DOI 查询工具 - 打包脚本
echo ========================================
echo.

echo [1/3] 正在安装依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 错误: 依赖安装失败！
    pause
    exit /b 1
)
echo.

echo [2/3] 正在打包应用程序...
pyinstaller --onefile --windowed --name "DOI查询工具" --clean doi_tool.py
if %errorlevel% neq 0 (
    echo 错误: 打包失败！
    pause
    exit /b 1
)
echo.

echo [3/3] 打包完成！
echo.
echo ========================================
echo  exe 文件位置: dist\DOI查询工具.exe
echo ========================================
echo.

pause


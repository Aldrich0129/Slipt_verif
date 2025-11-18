@echo off
REM PDF验证工具 - 打包脚本（Windows）

echo ======================================
echo PDF工资单验证工具 - 打包脚本
echo ======================================
echo.

REM 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python
    pause
    exit /b 1
)

echo 1. 创建虚拟环境...
python -m venv venv
call venv\Scripts\activate.bat

echo.
echo 2. 安装依赖...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo 3. 使用PyInstaller打包...
pyinstaller --onefile ^
    --windowed ^
    --name "PDF验证工具" ^
    --icon=NONE ^
    --add-data "requirements.txt;." ^
    pdf_validator.py

echo.
echo ======================================
echo 打包完成！
echo 可执行文件位于: dist\PDF验证工具.exe
echo ======================================
echo.

REM 清理
call venv\Scripts\deactivate.bat
pause

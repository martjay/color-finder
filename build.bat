@echo off
chcp 65001 >nul
echo ========================================
echo    颜色块识别器 - 构建EXE
echo ========================================
echo.

if not exist venv (
    echo [错误] 请先运行 install.bat
    pause
    exit /b 1
)

call venv\Scripts\activate.bat

echo [1/3] 安装 PyInstaller...
pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple

echo [2/3] 清理旧文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

echo [3/3] 构建EXE...
pyinstaller --onefile --windowed --name "颜色块识别器" --icon=icon.ico --add-data "icon.ico;." --clean app.py

if errorlevel 1 (
    echo [错误] 构建失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo    构建完成！
echo    EXE位于: dist\颜色块识别器.exe
echo ========================================
pause

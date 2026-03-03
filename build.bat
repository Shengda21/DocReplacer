@echo off
echo ========================================
echo   Building Document Replacer EXE
echo ========================================
pyinstaller --onefile --windowed --name "文档词汇替换工具" --icon=app_icon.ico --add-data "app_icon.ico;." --collect-data tkinterdnd2 main.py
echo ========================================
echo   Build complete! Check dist folder.
echo ========================================
pause

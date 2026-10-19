@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Building executable...
python build_exe.py

echo Build complete!
echo The executable is located in: dist\DataProcessingTool.exe
pause
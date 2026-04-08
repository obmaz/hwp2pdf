@echo off

echo [1/2] Installing required packages...
pip install -r requirements.txt

echo.
echo [2/2] Building EXE...
python -m PyInstaller --noconsole --onefile --windowed --collect-all tkinterdnd2 --name "HWP_to_PDF" src\main.py

echo.
echo Build complete! Check the 'dist' folder for 'HWP_to_PDF.exe'.
pause

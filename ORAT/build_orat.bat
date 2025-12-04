
@echo off
echo Building ORAT executable with icon...

:: Clean previous builds
rmdir /s /q build
rmdir /s /q dist
del /q ORAT.spec

:: Build with PyInstaller and custom icon
pyinstaller --noconsole --onefile --name ORAT --icon=orat_logo.ico main.py

echo Done! Find your executable in the dist\ folder as ORAT.exe
pause

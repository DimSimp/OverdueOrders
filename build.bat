@echo off
setlocal

set PYTHON=C:\Users\lorel\AppData\Local\Programs\Python\Python39\python.exe
set PIP=%PYTHON% -m pip
set PYINSTALLER=%PYTHON% -m PyInstaller

echo === Installing / upgrading PyInstaller ===
%PIP% install --upgrade pyinstaller
if errorlevel 1 (
    echo [ERROR] pip failed. Check that Python is installed correctly.
    pause
    exit /b 1
)

echo.
echo === Building Scarlett AIO ===
%PYINSTALLER% --clean ScarlettAIO.spec
if errorlevel 1 (
    echo [ERROR] Build failed. Check the output above for details.
    pause
    exit /b 1
)

echo.
echo === Build complete ===
echo Output: dist\Scarlett AIO\
echo.
pause

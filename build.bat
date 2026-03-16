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
echo === Reading version number ===
for /f "tokens=3 delims= " %%v in ('findstr "__version__" src\version.py') do set VERSION=%%~v
set VERSION=%VERSION:"=%
echo Version: %VERSION%

echo.
echo === Copying label settings into build output ===
if exist "data\shipping\label_settings.json" (
    if not exist "dist\Scarlett AIO\data\shipping" mkdir "dist\Scarlett AIO\data\shipping"
    copy /y "data\shipping\label_settings.json" "dist\Scarlett AIO\data\shipping\label_settings.json" >nul
    echo Copied label_settings.json to dist\Scarlett AIO\data\shipping\
) else (
    echo [INFO] No label_settings.json found -- build will use hardcoded defaults.
)

echo.
echo === Waiting for file locks to release ===
timeout /t 5 /nobreak >nul

echo.
echo === Creating release zip ===
set ZIPNAME=Scarlett AIO v%VERSION%.zip
if exist "%ZIPNAME%" del "%ZIPNAME%"
powershell -Command "Compress-Archive -Path 'dist\Scarlett AIO' -DestinationPath '%ZIPNAME%' -Force"
if errorlevel 1 (
    echo [WARNING] Zip creation failed -- distribute dist\Scarlett AIO\ manually.
) else (
    echo Created: %ZIPNAME%
)

echo.
echo === Deploying to server ===
set DEPLOY_PATH=\\SERVER\Project Folder\Order-Fulfillment-App\dist\Scarlett AIO
robocopy "dist\Scarlett AIO" "%DEPLOY_PATH%" /MIR /XF config.json /R:2 /W:3 >nul
if errorlevel 8 (
    echo [WARNING] Server deploy failed -- server may be offline. Distribute manually.
) else (
    echo Deployed to: %DEPLOY_PATH%
)

echo.
echo === Build complete ===
echo   Exe folder : dist\Scarlett AIO\
echo   Release zip: %ZIPNAME%
echo.
echo Next steps:
echo   1. Go to GitHub ^> Releases ^> Draft a new release
echo   2. Tag: v%VERSION%
echo   3. Attach "%ZIPNAME%" as a release asset
echo   4. Publish the release
echo.
pause
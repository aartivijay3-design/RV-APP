@echo off
echo ================================================
echo  BAWA Reiseverlauf Generator - Installation
echo ================================================
echo.

:: ── Find Python ────────────────────────────────────
set PYTHON=
for %%p in (
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%USERPROFILE%\AppData\Local\Programs\Python\Python313\python.exe"
    "%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe"
    "%USERPROFILE%\AppData\Local\Programs\Python\Python311\python.exe"
    "C:\ProgramData\Miniconda3\python.exe"
    "C:\ProgramData\Anaconda3\python.exe"
    "%USERPROFILE%\Miniconda3\python.exe"
    "%USERPROFILE%\anaconda3\python.exe"
) do (
    if exist %%p (
        set PYTHON=%%p
        goto :found_python
    )
)

:: Try PATH as last resort
python --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=python
    goto :found_python
)

echo ERROR: Python 3.10+ not found on this computer.
echo.
echo Please install Python from: https://www.python.org/downloads/
echo    - Download Python 3.12 (Windows installer 64-bit)
echo    - During install, CHECK the box "Add Python to PATH"
echo    - Then run this install.bat again
echo.
pause
exit /b 1

:found_python
echo [1/4] Python found: %PYTHON%
%PYTHON% --version

echo.
echo [2/4] Installing dependencies...
%PYTHON% -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: pip install failed. Check your internet connection.
    pause
    exit /b 1
)

echo.
echo [3/4] Checking template file...
if not exist "assets\template.docx" (
    echo WARNING: assets\template.docx not found.
    echo Copy the Baumgartner Japan reference template to: assets\template.docx
) else (
    echo Template found OK.
)

echo.
echo [4/4] Opening firewall port 8000 (so colleagues on the office network can reach the app)...
netsh advfirewall firewall show rule name="BAWA Reiseverlauf Generator (port 8000)" >nul 2>&1
if errorlevel 1 (
    netsh advfirewall firewall add rule name="BAWA Reiseverlauf Generator (port 8000)" dir=in action=allow protocol=TCP localport=8000 profile=any >nul 2>&1
    if errorlevel 1 (
        echo WARNING: Could not add the firewall rule ^(this step needs Administrator rights^).
        echo Right-click install.bat and choose "Run as administrator", then run it again.
        echo Without this, the app will work on this computer but NOT for colleagues on the network.
    ) else (
        echo Firewall rule added — colleagues on the office network can now reach this app.
    )
) else (
    echo Firewall rule already present.
)

:: Save the Python path for start.bat
echo %PYTHON%> .python_path.txt

echo.
echo ================================================
echo  Installation complete!
echo.
echo  NEXT STEP: Open .env in Notepad and add your
echo  Gemini API key:
echo     GEMINI_API_KEY=your-key-here
echo.
echo  Get your key (free tier available) at:
echo     https://aistudio.google.com/apikey
echo.
echo  Then run start.bat to launch the server.
echo ================================================
pause

@echo off
echo ================================================
echo  BAWA Reiseverlauf Generator
echo ================================================
echo.

:: ── Find Python ────────────────────────────────────
if exist .python_path.txt (
    set /p PYTHON=<.python_path.txt
) else (
    set PYTHON=python
)

:: Validate Python works
"%PYTHON%" --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please run install.bat first.
    pause
    exit /b 1
)

:: ── Check API key ───────────────────────────────────
if not exist .env (
    echo ERROR: .env file not found.
    echo Create a .env file with: ANTHROPIC_API_KEY=sk-ant-...
    pause
    exit /b 1
)

findstr /C:"your-api-key-here" .env >nul 2>&1
if not errorlevel 1 (
    echo ERROR: API key not set in .env file.
    echo Open .env with Notepad and replace "your-api-key-here" with your key.
    echo Get your FREE key from: https://aistudio.google.com/apikey
    echo.
    pause
    exit /b 1
)

:: ── Get local IP ────────────────────────────────────
set LOCAL_IP=
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /R /C:"IPv4.*[0-9]"') do (
    if not defined LOCAL_IP (
        set LOCAL_IP=%%a
    )
)
set LOCAL_IP=%LOCAL_IP: =%

:: ── Launch ─────────────────────────────────────────
echo  Server starting...
echo.
echo  ┌─────────────────────────────────────────┐
echo  │  Open in browser:                       │
echo  │                                         │
echo  │  Your computer:  http://localhost:8000  │
if defined LOCAL_IP (
echo  │  Colleagues:     http://%LOCAL_IP%:8000
)
echo  │                                         │
echo  │  Share the Colleagues link with anyone  │
echo  │  on the same office network.            │
echo  └─────────────────────────────────────────┘
echo.
echo  Press Ctrl+C to stop the server.
echo.

"%PYTHON%" -m uvicorn app:app --host 0.0.0.0 --port 8000

pause

@echo off
echo Testing BAWA Reiseverlauf Generator...
echo.
curl -s http://localhost:8000/health
echo.
echo.
echo If you see {"status":"ok","template":true} the server is running correctly.
echo If you see an error, make sure start.bat is running first.
pause

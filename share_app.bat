@echo off
echo Starting Cloudflare Tunnel...
echo This will generate a public link (e.g. https://crazy-name.trycloudflare.com).
echo Copy that link and send it to your client.
echo.
echo NOTE: Keep this window OPEN while they are testing.
echo.
cloudflared.exe tunnel --url http://127.0.0.1:5000
pause

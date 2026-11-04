@echo off
title Albrost Accounts Online
cd /d "%~dp0"
start "Albrost server" python server.py
echo.
echo Look below for a line with https://....trycloudflare.com - that is the link.
echo Keep this window open. Closing it stops the link.
echo.
cloudflared tunnel --no-autoupdate --url http://127.0.0.1:8766

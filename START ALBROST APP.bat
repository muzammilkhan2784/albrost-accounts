@echo off
title Albrost Accounts
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0app\server.ps1"
exit

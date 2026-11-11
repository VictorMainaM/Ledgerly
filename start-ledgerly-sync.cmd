@echo off
title Ledgerly Excel Sync
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0ledgerly-sync.ps1"
if errorlevel 1 pause

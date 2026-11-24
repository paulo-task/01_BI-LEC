@echo off
title Conectar VPN Engelmig
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0Conectar_VPN.ps1"

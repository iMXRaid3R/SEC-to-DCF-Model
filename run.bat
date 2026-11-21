@echo off
rem Launcher for Windows. cmd:  run NVDA   PowerShell:  .\run.bat NVDA   or double-click for the menu
cd /d "%~dp0"
uv run --with requests --with pandas --with openpyxl --with pywin32 python edgar_etl.py %*

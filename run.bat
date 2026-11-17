@echo off
rem Launcher for cmd / PowerShell:  run ISRG   ^|   run --list holdings   ^|   run --rollup holdings
cd /d "%~dp0"
uv run --with requests --with pandas --with openpyxl --with pywin32 python edgar_etl.py %*

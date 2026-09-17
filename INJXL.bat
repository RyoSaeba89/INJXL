@echo off
REM Lance INJXL sans fenetre de console (double-clic).
cd /d "%~dp0"
start "" pythonw INJXL.py %*

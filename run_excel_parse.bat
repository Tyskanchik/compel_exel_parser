@echo off

echo Start script

cd %~dp0

call .venv\Scripts\activate.bat
call python main.py

echo Complete
pause
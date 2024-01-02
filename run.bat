@echo off

rem Define variables
set VENV_NAME=venv
set MAIN_SCRIPT=main.py

rem Activate virtual environment
call %VENV_NAME%\Scripts\activate

rem Run main.py
python %MAIN_SCRIPT%

rem Deactivate virtual environment
deactivate

echo Script execution completed.

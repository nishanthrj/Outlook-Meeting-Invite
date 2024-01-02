@echo off
setlocal enabledelayedexpansion

rem Define variables
set VENV_NAME=venv
set REQUIREMENTS_FILE=requirements.txt

rem Verify Python installation
python --version

rem Set up virtual environment
echo Creating virtual environment %VENV_NAME%...
python -m venv %VENV_NAME%

rem Activate virtual environment
call %VENV_NAME%\Scripts\activate

rem Install packages from requirements file
echo Installing packages from %REQUIREMENTS_FILE%...
pip install -r %REQUIREMENTS_FILE%

echo Setup completed successfully.

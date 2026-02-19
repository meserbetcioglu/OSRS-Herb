@echo off
setlocal

set "ROOT=%~dp0"
set "VENV=%ROOT%.venv"

if not exist "%VENV%" (
  echo Creating venv...
  py -3 -m venv "%VENV%"
)

call "%VENV%\Scripts\activate.bat"

python -m pip install --upgrade pip
python -m pip install -r "%ROOT%requirements.txt"

echo.
echo Setup complete.
echo Activate with: %VENV%\Scripts\activate.bat
endlocal

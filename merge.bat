@echo off
:: check Python environment
python --version 2>NUL
if not errorlevel 0 goto PythonNotInstalledError
:: check and update pip3
pip3 install --upgrade pip
:: check dependencies
pip3 install --ignore-installed -r requirements.txt
cls
echo Please select where you want to save the merged file:
CALL :WAIT 1
for /f "tokens=*" %%i in ('cscript //nologo bin\browse.vbs') do set TF=%%i
cls
:: run merge_pdf
python %cd%\merge_pdf.py %TF%
goto End

:PythonNotInstalledError
echo Python 3 is not installed, please visit https://www.anaconda.com/distribution/ to download latest Python 3.x

:End
pause

:wait
PING 127.0.0.1 -w 1 -n %~1 2>NUL>NUL
GOTO :EOF
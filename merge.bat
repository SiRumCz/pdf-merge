@echo off
:: check Python environment
python --version 2>NUL
if not errorlevel 0 goto PythonNotInstalledError
:: check and update pip3
pip3 install --upgrade pip
:: check dependencies
pip3 install --ignore-installed -r requirements.txt 
:: run merge_pdf
python %cd%\merge_pdf.py %*
goto End
:PythonNotInstalledError
echo Python 3 is not installed, please visit https://www.anaconda.com/distribution/ to download latest Python 3.x
:End
pause
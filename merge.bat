:: check update pip3
pip3 install --upgrade pip
:: check update dependencies
pip3 install --ignore-installed -r requirements.txt 
:: run merge_pdf
python "D:\myGithubRepo\pdf-merge\merge_pdf.py" %*
pause
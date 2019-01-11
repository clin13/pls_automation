# pls_automation
Steps:

# 1a. Python 2
   - Install Python 2.x (Ex: Python 2.7.9)
   - Install xlwt 0.7.5
   - https://pypi.python.org/pypi/xlwt
   Description:
   - Library to create spreadsheet files compatible with MS Excel 
   97/2000/XP/2003 XLS files, on any platform, with Python 2.3 to 2.7

   - Extract xlwt-0.7.5.tar.gz to C:\Python27\Lib\site-packages\xlwt-0.7.5
   - Open command line prompt
   - cd C:\Python27\Lib\site-packages\xlwt-0.7.5
   - $python setup.py install

# 1b. Python 3
   - Install Python 3.7 (Ex: Python 3.7.2)
   - Install xlwt - $pip install xlwt

# 2. Create the following directories and files
   - a. Root directory - Date__Component_Lists.
   - b. There are all product folders and pls_by_lang folder.
   - c. Server.ini are in product folder.
   - d. Component lists are in the folder by language in product folder.
   - e. CSV files with PLM/PLS version are in the pls_by_lang folder.
   - f. A log file called pls.log in the root directory.
   
# 3. Generate result CSV file
   1. An Excel file "PLM_PLS_List.xls" in the root directory. Each worksheet, by language, collects PLM and PLS version by products.

# Note: 
   1. Please don't modify server.ini while dowloading all files.
   2. If AU server is updating, we need to wait until server.ini is ready to use.
   3. Current delay time - Server.ini: ~1 second, Component list: ~2 second


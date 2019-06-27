QUICK START GUIDE

DEPENDENCIES:
Python 3.x
pip package manager

pip packages:
	openpyxl
	datetime
	
SETUP:

1. Verify python is installed
	in cmd, type "python" for version details
	if not installed, download from 'https://www.python.org/downloads/windows/'

2. Install pip
	save text from 'https://bootstrap.pypa.io/get-pip.py' as get-pip.py
	in cmd, navigate directory with "cd" command to folder containing get-pip.py
	type, "python get-pip.py"
	verify installation with "pip --version"
	optional: upgrade to latest version with "python -m pip install --upgrade pip"

3. Install pip packages
	in cmd, type:
	"pip install openpyxl"
	"pip install datetime"

USAGE:

1. Open .csv and rename first sheet to "sheet1"
2. Save .csv as an .xlsx workbook named "ngt_log"
3. Verify that ngt_log.xlsx and parseData.py are in the same folder
4. Double click script to run. Script window will close by itself when finished.
5. Graphs will be located in the "Data Analysis" sheet of ngt_log and data will be located in "log.txt"






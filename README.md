## excel dateparser
parse dates from Excel spreadsheet from one language to another using 'dateparser'

Assuming you have Python 3.x and have pip installed, ensure 'virtualenv' is installed, because we will use it to make a virtual environment for this repository.

`pip install virtualenv`

Clone this repo to a local directory.

`C:\code> git clone https://github.com/jmasselink/excel_dateparser.git`

Create a Python virtual environment called **excel_dateparser** where you store your necessary libraries.
 - in Windows, using command line:

`C:\Python37>python -m virtualenv 'C:\venv\excel_dateparser'`

Navigate to C:\venv\excel_dateparser and activate the environment

`cd C:\venv\excel_dateparser`
`C:\venv\excel_dateparser>Scripts\activate`

Install required libraries using either the requirements file:

`(excel_dateparser) C:\venv\excel_dateparser\Scripts> pip install -r [PATH]/requirements_python.txt`

  or using Pip commands:

  `pip install pandas`

  `pip install dateparser`

  `pip install openpyxl`

(optional) if you want to use jupyter lab:

>run 3 pip install commands above plus the following:

`pip install jupyterlab`

>OR:
>run pip install using a different requirements file.

`(excel_dateparser) C:\venv\excel_dateparser\Scripts> pip install -r [PATH]/requirements_jupyter.txt`

  To launch jupyter lab:

`(pyexcel) C:\code\excel_dateparser> jupyter lab`

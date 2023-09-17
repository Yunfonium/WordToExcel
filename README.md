# WordToExcel

Before starting transforming Docx file to Xsl file, you should install Python or Poetry.

Refenece:

Python: https://www.python.org/downloads/

Poetry: https://python-poetry.org/docs/


## Dependencies

python = "^3.8" 

python-docx = "^0.8.11"

pandas = "1.1.5"


## Install Dependencies

If use pip, just execute this command:

    pip install requirements.txt


If use Poetry, execute this command to create venv:

    poetry install

and then start the venv:

    poetry shell


## Start transform
Execute this command on the root path of this repo

    python start.py


And start.py will transform the "./word_data_身份資料文件.docx" to "~/excel_data/tmp.xlsx" with the assigned table format.
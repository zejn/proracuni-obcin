
all:
	(. venv/bin/activate && python xls2csv.py csv)

setup:
	virtualenv venv
	(. venv/bin/activate && pip install -r requirements.txt)


.PHONY: run lint format check

install:
	python -m pip install --upgrade pip
	python -m pip install -r requirements.txt

run:
	PYTHONPATH=. streamlit run src/app/app.py --server.headless true

lint:
	ruff check .
	pyright .

format:
	ruff format .

check:
	ruff check . && ruff format --check .
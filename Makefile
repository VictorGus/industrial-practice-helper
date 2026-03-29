VENV := .venv
PYTHON := $(VENV)/bin/python

.PHONY: venv install run-tg run-max lint test clean

venv:
	python3 -m venv $(VENV)

install: venv
	$(VENV)/bin/pip install -e ".[dev]"

run-tg: install
	$(PYTHON) -m bot_tg.main

run-max: install
	$(PYTHON) -m bot_max.main

lint:
	$(VENV)/bin/ruff check .

test:
	$(VENV)/bin/pytest -v

build: install
	$(VENV)/bin/pyinstaller --onefile --name practice-bot \
		--hidden-import=common \
		--hidden-import=common.config \
		--hidden-import=common.storage \
		--hidden-import=common.logger \
		--hidden-import=bot_tg \
		--hidden-import=bot_tg.handlers \
		--hidden-import=openpyxl \
		bot_tg/main.py

docker-build:
	docker build -t practice-bot .

docker-run:
	docker run --rm --env-file .env practice-bot

clean:
	rm -rf $(VENV) *.egg-info __pycache__ .pytest_cache .ruff_cache

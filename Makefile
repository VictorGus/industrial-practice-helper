VENV := .venv
PYTHON := $(VENV)/bin/python

.PHONY: venv install run-tg run-max lint test clean \
       service-install service-uninstall service-start service-stop service-restart service-status service-logs

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

# --- systemd service ---
SERVICE_NAME := practice-bot
SERVICE_FILE := /etc/systemd/system/$(SERVICE_NAME).service
PROJECT_DIR := $(shell pwd)

define SERVICE_UNIT
[Unit]
Description=Practice Bot (Telegram)
After=network.target

[Service]
Type=simple
User=$(USER)
WorkingDirectory=$(PROJECT_DIR)
EnvironmentFile=$(PROJECT_DIR)/.env
ExecStart=$(PROJECT_DIR)/$(PYTHON) -m bot_tg.main
Restart=on-failure
RestartSec=5

[Install]
WantedBy=multi-user.target
endef
export SERVICE_UNIT

service-install: install
	echo "$$SERVICE_UNIT" | sudo tee $(SERVICE_FILE) > /dev/null
	sudo systemctl daemon-reload
	sudo systemctl enable $(SERVICE_NAME)

service-uninstall:
	sudo systemctl disable $(SERVICE_NAME) || true
	sudo systemctl stop $(SERVICE_NAME) || true
	sudo rm -f $(SERVICE_FILE)
	sudo systemctl daemon-reload

service-start:
	sudo systemctl start $(SERVICE_NAME)

service-stop:
	sudo systemctl stop $(SERVICE_NAME)

service-restart:
	sudo systemctl restart $(SERVICE_NAME)

service-status:
	sudo systemctl status $(SERVICE_NAME)

service-logs:
	journalctl -u $(SERVICE_NAME) -f

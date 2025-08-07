.PHONY: help install run test clean build

# The first target is the default one when running 'make' without arguments.
help:
	@echo "Available targets:"
	@echo "  install - Install dependencies from requirements.txt"
	@echo "  run     - Run the main application UI"
	@echo "  test    - Run the test suite"
	@echo "  clean - Remove cache and other generated files"

install:
	python -m uv pip install -r requirements.txt

run:
	python -m cuttingstock.ui

test:
	python -m pytest


build:
	pyinstaller --log-level DEBUG --clean --noconfirm --windowed \
	    --collect-data pulp \
	    --hidden-import scipy \
	    --hidden-import scipy._cyutility \
	    --collect-all xgboost \
	    --collect-all connectorx \
	    --add-data "./model;model" \
	    --name order-optimizer cuttingstock\ui.py

clean:
	# Note: These commands use Unix-style tools ('rm', 'find'). On Windows,
	# you may need to run this from a shell like Git Bash which includes these tools.
	-rm -rf cache
	-rm -rf .pytest_cache
	-find . -type d -name "__pycache__" -exec rm -rf {} +
	-find . -type f -name "*.pyc" -delete


pyinstaller --clean --noconfirm --windowed --collect-data pulp --name order-optimizer ui.py


pyinstaller --clean --noconfirm --windowed --onefile --collect-data pulp --name order-optimizer-simple simple_ui.py

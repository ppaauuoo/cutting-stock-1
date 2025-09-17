#!/usr/bin/env python3
"""
CLI wrapper for the cutting stock optimizer application.
Supports --factory argument to auto-select factory and run main algorithm.
"""

import argparse
import sys
import os
import signal
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer, QCoreApplication

# Add the cuttingstock module to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cuttingstock.ui import CuttingOptimizerUI


def signal_handler(signum, frame):
    """Handle system signals for graceful shutdown"""
    print(f"\nReceived signal {signum}, shutting down gracefully...")
    QCoreApplication.quit()


def main():
    parser = argparse.ArgumentParser(description='Cutting Stock Optimizer CLI')
    parser.add_argument(
        '--factory', 
        type=int, 
        choices=[1, 2, 3, 4, 5],
        help='Factory number to select (1-5). If provided, will auto-run the main algorithm.'
    )
    parser.add_argument(
        '--no-auto-close',
        action='store_true',
        help='Prevent automatic closing after export (keep GUI open)'
    )
    
    args = parser.parse_args()
    
    # Set up signal handlers for graceful shutdown
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    
    # Set up environment variables for Windows
    if sys.platform == "win32":
        os.environ["QT_QPA_PLATFORM"] = "windows:fontengine=freetype"
        os.environ["PYTHONIOENCODING"] = "utf-8"
        # Disable Windows file locking to prevent hanging
        os.environ["QT_FILE_LOCKING"] = "0"
    
    # Create QApplication
    app = QApplication(sys.argv)
    
    # Set up proper cleanup on exit
    def handle_quit():
        """Handle application quit"""
        # Force immediate cleanup and quit
        from PyQt5.QtWidgets import QApplication
        QApplication.processEvents()
        QCoreApplication.quit()

    # Connect quit signal
    app.aboutToQuit.connect(handle_quit)
    
    # Create main window with auto-export and auto-close enabled when factory is specified
    auto_export = args.factory is not None
    auto_close = args.factory is not None and not args.no_auto_close
    window = CuttingOptimizerUI(auto_export=auto_export, auto_close=auto_close)
    window.show()
    
    # If factory argument is provided, set the factory and auto-run
    if args.factory is not None:
        factory_str = str(args.factory)
        
        # Use QTimer to ensure the UI is fully initialized before running the algorithm
        def auto_run():
            # Set the factory combobox
            index = window.factory_combo.findText(factory_str)
            if index >= 0:
                window.factory_combo.setCurrentIndex(index)
                window.log_message(f"🏭 Auto-selected factory: {factory_str}")
                
                # Auto-start the main algorithm
                window.log_message("🚀 Auto-starting main algorithm...")
                QTimer.singleShot(100, window.start_main_loop)
            else:
                window.log_message(f"❌ Factory {factory_str} not found in combobox")
        
        # Schedule auto-run after UI is ready
        QTimer.singleShot(500, auto_run)
    
    # Start the application event loop
    try:
        exit_code = app.exec_()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\nApplication interrupted by user")
        sys.exit(0)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
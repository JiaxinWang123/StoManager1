"""
StoManager1 - Stomata Detection and Analysis Tool

Main entry point for the application.

Author: Jiaxin Wang
Contact: jiaxinwang@cornell.edu; jiaxinwang362@gmail.com
"""

import sys
import torch
from PyQt5 import QtWidgets
from ui.main_window import MainWindow


def main():
    """Initialize and run the application."""
    app = QtWidgets.QApplication(sys.argv)
    
    # Set application metadata
    app.setApplicationName("StoManager1")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("Jiaxin Wang")
    
    # Create and show main window
    window = MainWindow()
    window.show()
    
    # Start event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

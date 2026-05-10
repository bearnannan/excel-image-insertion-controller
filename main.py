# =================================================================
# Project: Excel Image Insertion Controller
# Author: WATCHARA MANADEE
# Date: 2026
# License: MIT
# Description: Main entry point for the standalone application.
# =================================================================

import sys
import os
import customtkinter as ctk

from excel.excel_controller import ExcelController
from gui.main_window import MainWindow

def main():
    # Set appearance mode and color theme
    ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
    ctk.set_default_color_theme("green")  # Themes: "blue" (standard), "green", "dark-blue"
    
    # Initialize Controller
    controller = ExcelController()
    
    # Initialize GUI
    app = MainWindow(controller)
    
    # Wire GUI callback to Controller
    controller.gui_callback = app.handle_event
    
    # Set App Icon if exists
    try:
        # Check if running as PyInstaller bundle
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, 'icon.ico')
        else:
            icon_path = 'icon.ico' # We'll put it in root for simplicity
            
        if os.path.exists(icon_path):
            app.iconbitmap(icon_path)
    except Exception as e:
        print(f"Icon warning: {e}")

    # Open Floating Toolbar by default
    app._toggle_toolbar()

    # Start Controller Thread
    controller.start_monitoring()

    # Handle Closing
    app.protocol("WM_DELETE_WINDOW", app.on_closing)

    # Start GUI Loop
    app.mainloop()

if __name__ == "__main__":
    main()

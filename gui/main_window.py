# =================================================================
# Project: Excel Image Insertion Controller
# Component: Main GUI Dashboard
# Author: WATCHARA MANADEE
# License: MIT
# =================================================================

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog
from plyer import notification
import threading

from gui.floating_toolbar import FloatingToolbar

# We must use TkinterDnD.Tk as the root to support drag and drop system-wide in Tkinter
class CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

class MainWindow(CTk):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        
        self.title("Excel Image Insertion Controller")
        self.geometry("400x500")
        
        self.toolbar = None
        self.show_toolbar = ctk.BooleanVar(value=True)
        
        # Setup UI
        self._build_ui()

    def _build_ui(self):
        # Header
        self.header = ctk.CTkLabel(self, text="Excel Image Tool", font=ctk.CTkFont(size=20, weight="bold"))
        self.header.pack(pady=(20, 10))

        # Status Panel
        self.status_frame = ctk.CTkFrame(self)
        self.status_frame.pack(fill="x", padx=20, pady=10)
        
        self.excel_status = ctk.CTkLabel(self.status_frame, text="Excel: Not Connected", text_color="#FF5252", font=ctk.CTkFont(weight="bold"))
        self.excel_status.pack(pady=5)
        
        self.mode_status = ctk.CTkLabel(self.status_frame, text="Image Mode: OFF")
        self.mode_status.pack(pady=5)

        # Controls
        self.toggle_mode_btn = ctk.CTkButton(self, text="Enable Image Mode", command=self.controller.toggle_image_mode)
        self.toggle_mode_btn.pack(pady=10)

        # Settings
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.pack(fill="x", padx=20, pady=10)
        
        self.auto_fit_var = ctk.BooleanVar(value=False)
        self.auto_center_var = ctk.BooleanVar(value=True)
        
        ctk.CTkSwitch(self.settings_frame, text="Auto-fit Row Height", variable=self.auto_fit_var, command=self._update_settings).pack(pady=5, padx=10, anchor="w")
        ctk.CTkSwitch(self.settings_frame, text="Auto Center Image", variable=self.auto_center_var, command=self._update_settings).pack(pady=5, padx=10, anchor="w")

        # Toolbar Toggle
        ctk.CTkSwitch(self, text="Show Floating Toolbar", variable=self.show_toolbar, command=self._toggle_toolbar).pack(pady=10)

    def _update_settings(self):
        self.controller.set_settings({
            'auto_fit_row': self.auto_fit_var.get(),
            'auto_center': self.auto_center_var.get()
        })

    def _toggle_toolbar(self):
        if self.show_toolbar.get():
            if self.toolbar:
                self.toolbar.deiconify()
            else:
                self.toolbar = FloatingToolbar(self, self.controller)
        else:
            if self.toolbar:
                self.toolbar.withdraw()

    def handle_event(self, event_type, data):
        """Thread-safe event routing called by the controller."""
        # CustomTkinter .after() is thread-safe for basic UI updates
        if event_type == "connected":
            self.after(0, self._update_connection_status, data)
        elif event_type == "mode_changed":
            self.after(0, self._update_mode_status, data)
        elif event_type == "request_file_dialog":
            self.after(0, self._open_file_dialog, data["sheet_name"], data["address"])
        elif event_type == "images_inserted":
            self.after(0, self._show_notification, f"Inserted {data} image(s) successfully.", "Success")
        elif event_type == "error":
            self.after(0, self._show_notification, data, "Error")

    def _update_connection_status(self, connected):
        if connected:
            self.excel_status.configure(text="Excel: Connected", text_color="#217346")
            self._show_notification("Connected to Excel.", "Status")
        else:
            self.excel_status.configure(text="Excel: Not Connected", text_color="#FF5252")

    def _update_mode_status(self, is_on):
        if is_on:
            self.mode_status.configure(text="Image Mode: ON", text_color="#FFC107")
            self.toggle_mode_btn.configure(text="Disable Image Mode")
            self._show_notification("Image Mode Enabled.", "Status")
        else:
            self.mode_status.configure(text="Image Mode: OFF", text_color="white") # Or default text color
            self.toggle_mode_btn.configure(text="Enable Image Mode")
        
        if self.toolbar:
            self.toolbar.update_mode_status(is_on)

    def _open_file_dialog(self, sheet_name, target_address):
        # We must ensure this dialog stays on top
        file_paths = filedialog.askopenfilenames(
            title="Select Images to Insert",
            filetypes=[("Images", "*.webp;*.jpg;*.jpeg;*.png;*.bmp")]
        )
        if file_paths:
            # We process them in a background thread to not block the GUI or COM
            threading.Thread(
                target=self.controller.process_selected_files, 
                args=(sheet_name, target_address, file_paths),
                daemon=True
            ).start()

    def _show_notification(self, message, title):
        try:
            notification.notify(
                title=f"Excel Tool - {title}",
                message=message,
                app_name="Excel Image Tool",
                timeout=3
            )
        except Exception as e:
            print(f"Notification error: {e}")

    def on_closing(self):
        self.controller.stop_monitoring()
        self.destroy()

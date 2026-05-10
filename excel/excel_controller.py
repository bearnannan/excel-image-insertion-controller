# =================================================================
# Project: Excel Image Insertion Controller
# Component: Excel Controller
# Author: WATCHARA MANADEE
# License: MIT
# =================================================================

import win32com.client
import pythoncom
import threading
import time
from hooks.event_bridge import ExcelEventBridge
from excel.image_handler import ImageHandler

class ExcelController:
    def __init__(self, gui_callback=None):
        self.gui_callback = gui_callback
        self.image_handler = ImageHandler()
        
        self.is_image_mode = False
        self.is_connected = False
        
        self.excel_app = None
        self.event_bridge = None
        self.monitor_thread = None
        self.running = False
        self.settings = {
            'auto_fit_row': False,
            'auto_center': True,
            'padding': 2
        }

    def start_monitoring(self):
        self.running = True
        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()

    def stop_monitoring(self):
        self.running = False
        self.image_handler.cleanup()

    def set_settings(self, settings_dict):
        self.settings.update(settings_dict)

    def is_image_mode_enabled(self):
        return self.is_image_mode

    def set_image_mode(self, state):
        self.is_image_mode = state
        self._notify_gui("mode_changed", state)

    def toggle_image_mode(self):
        self.set_image_mode(not self.is_image_mode)

    def _notify_gui(self, event_type, data=None):
        if self.gui_callback:
            # We must use thread-safe calls if the GUI framework requires it.
            # In Tkinter/CustomTkinter, we usually use `after` or similar.
            # The gui_callback should handle this thread dispatch.
            self.gui_callback(event_type, data)

    def _monitor_loop(self):
        # Set the controller reference in the bridge class
        ExcelEventBridge.controller = self
        
        while self.running:
            pythoncom.CoInitialize()
            try:
                # Try to connect to an existing Excel instance
                self.excel_app = win32com.client.GetActiveObject("Excel.Application")
                
                if not self.is_connected:
                    self.is_connected = True
                    self._notify_gui("connected", True)
                
                # Bind events
                self.event_bridge = win32com.client.WithEvents(self.excel_app, ExcelEventBridge)
                
                # Event pump loop
                while self.running and self.is_connected:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.05)
                    
                    # Check if connection is still alive
                    try:
                        _ = self.excel_app.Name
                    except Exception:
                        self.is_connected = False
                        self._notify_gui("connected", False)
                        break
                        
            except Exception:
                if self.is_connected:
                    self.is_connected = False
                    self._notify_gui("connected", False)
                
                # Wait before retrying
                time.sleep(2.0)
            finally:
                pythoncom.CoUninitialize()

    def handle_image_insertion(self, sheet, target_range):
        """Called by event bridge when right click happens."""
        try:
            # We pass identifiers instead of COM objects to avoid thread-safety issues
            data = {
                "sheet_name": sheet.Name,
                "address": target_range.Address
            }
            self._notify_gui("request_file_dialog", data)
        except Exception as e:
            self._notify_gui("error", f"Event context error: {e}")

    def process_selected_files(self, sheet_name, target_address, file_paths):
        """Called by GUI after files are selected. Runs in its own thread."""
        if not file_paths:
            return

        pythoncom.CoInitialize()
        try:
            # Re-acquire Excel objects in this thread
            excel = win32com.client.GetActiveObject("Excel.Application")
            sheet = excel.ActiveSheet # Assume the sheet is still active
            
            # If the user switched sheets, we try to find the right one
            if sheet.Name != sheet_name:
                try:
                    sheet = excel.ActiveWorkbook.Sheets(sheet_name)
                except:
                    pass # Fallback to active sheet
                    
            target_range = sheet.Range(target_address)
            
            success_count = 0
            current_target = target_range

            for path in file_paths:
                if self.image_handler.insert_image(sheet, current_target, path, self.settings):
                    success_count += 1
                    # Move target down by one row for batch insertion
                    try:
                        current_target = current_target.Offset(1, 0)
                    except:
                        break
                    
            if success_count > 0:
                self._notify_gui("images_inserted", success_count)
            else:
                self._notify_gui("error", "Failed to insert images. Check if Excel is busy.")
                
        except Exception as e:
            self._notify_gui("error", f"Insertion thread error: {e}")
        finally:
            pythoncom.CoUninitialize()

    def insert_dragged_file(self, file_path):
        """Called by Floating Toolbar when a file is dropped."""
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            sheet = excel.ActiveSheet
            target_range = excel.ActiveCell
            
            if self.image_handler.insert_image(sheet, target_range, file_path, self.settings):
                self._notify_gui("images_inserted", 1)
            else:
                self._notify_gui("error", "Drag & Drop failed. Excel might be in edit mode.")
        except Exception as e:
            self._notify_gui("error", f"Drag & Drop insertion failed: {e}")
        finally:
            pythoncom.CoUninitialize()

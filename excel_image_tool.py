import os
import time
import threading
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
from PIL import Image, ImageTk
import pystray
from pystray import MenuItem as item

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- CONFIGURATION ---
TOGGLE_SPEED = 1.0
IMAGE_PADDING = 2
WINDOW_TITLE = "Excel Image Tool"
PRIMARY_COLOR = "#217346"
ACCENT_COLOR = "#FFC107"
BG_DARK = "#1E1E1E"
ICON_PATH = resource_path("icon.png")

class ExcelImageToolUI:
    def __init__(self):
        self.is_image_mode = False
        self.last_click_time = 0
        self.is_connected = False
        
        # Setup Main UI Window
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)
        self.root.geometry("280x80+100+100")
        self.root.overrideredirect(True) 
        self.root.attributes("-topmost", True)
        self.root.config(bg=BG_DARK)
        
        # Make it draggable
        self.root.bind("<Button-1>", self.start_move)
        self.root.bind("<B1-Motion>", self.do_move)
        
        # UI Elements
        self.main_frame = tk.Frame(self.root, bg=BG_DARK, highlightthickness=1, highlightbackground=PRIMARY_COLOR)
        self.main_frame.pack(fill="both", expand=True)
        
        self.status_label = tk.Label(
            self.main_frame, text="CONNECTING...", fg="#AAAAAA", bg=BG_DARK,
            font=("Segoe UI", 12, "bold")
        )
        self.status_label.pack(pady=(15, 5))
        
        self.hint_label = tk.Label(
            self.main_frame, text="Please open Excel", fg="#666666", bg=BG_DARK,
            font=("Segoe UI", 8)
        )
        self.hint_label.pack()

        # Controls
        self.close_btn = tk.Button(
            self.main_frame, text="—", fg="white", bg=BG_DARK, 
            bd=0, font=("Arial", 12), command=self.hide_window
        )
        self.close_btn.place(x=255, y=2)

        self.tray_icon = None
        self.setup_tray()

    def start_move(self, event):
        self.x, self.y = event.x, event.y

    def do_move(self, event):
        x = self.root.winfo_x() + (event.x - self.x)
        y = self.root.winfo_y() + (event.y - self.y)
        self.root.geometry(f"+{x}+{y}")

    def update_connection_status(self, connected):
        self.is_connected = connected
        if not connected:
            self.status_label.config(text="EXCEL NOT FOUND", fg="#FF5252")
            self.hint_label.config(text="Waiting for Excel...")
            self.main_frame.config(highlightbackground="#FF5252")
        else:
            self.update_status()

    def update_status(self):
        if not self.is_connected: return
        if self.is_image_mode:
            self.status_label.config(text="IMAGE MODE: ON", fg=ACCENT_COLOR)
            self.hint_label.config(text="Right-click to insert")
            self.main_frame.config(highlightbackground=ACCENT_COLOR)
            if self.tray_icon: self.tray_icon.title = "Excel Image Tool: ON"
        else:
            self.status_label.config(text="IMAGE MODE: OFF", fg="white")
            self.hint_label.config(text="Double-click 2x to toggle")
            self.main_frame.config(highlightbackground=PRIMARY_COLOR)
            if self.tray_icon: self.tray_icon.title = "Excel Image Tool: OFF"

    def toggle_mode(self, icon=None, item=None):
        self.is_image_mode = not self.is_image_mode
        self.root.after(0, self.update_status)

    def OnSheetBeforeDoubleClick(self, Sh, Target, Cancel):
        current_time = time.time()
        if current_time - self.last_click_time < TOGGLE_SPEED and self.last_click_time > 0:
            self.toggle_mode()
            self.last_click_time = 0
            return True
        else:
            self.last_click_time = current_time
            if self.is_image_mode: return True
        return False

    def OnSheetBeforeRightClick(self, Sh, Target, Cancel):
        if self.is_image_mode:
            self.root.after(0, self.pick_and_insert, Sh, Target)
            return True
        return False

    def pick_and_insert(self, Sh, Target):
        # We need a new root for dialog because the main one is often withdrawn/sticky
        temp_root = tk.Tk()
        temp_root.withdraw()
        file_path = filedialog.askopenfilename(
            parent=temp_root,
            title=f"Select Image for {Target.Address}",
            filetypes=[("Images", "*.webp *.jpg *.jpeg *.png *.bmp")]
        )
        temp_root.destroy()
        
        if file_path:
            try:
                for shape in Sh.Shapes:
                    if shape.TopLeftCell.Address == Target.Address: shape.Delete()
                img = Sh.Shapes.AddPicture(file_path, False, True, Target.Left, Target.Top, -1, -1)
                img.LockAspectRatio = False
                img.Left, img.Top = Target.Left + 1, Target.Top + 1
                img.Width, img.Height = Target.Width - 2, Target.Height - 2
                img.Placement = 1
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def setup_tray(self):
        try:
            image = Image.open(ICON_PATH)
        except:
            image = Image.new('RGB', (64, 64), color=PRIMARY_COLOR)

        menu = (
            item('Toggle Mode', self.toggle_mode),
            item('Show/Hide Window', self.toggle_window),
            item('Exit', self.exit_app)
        )
        self.tray_icon = pystray.Icon("ExcelImageTool", image, "Excel Image Tool", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def toggle_window(self, icon=None, item=None):
        if self.root.winfo_viewable():
            self.root.withdraw()
        else:
            self.root.deiconify()
            self.root.attributes("-topmost", True)

    def hide_window(self):
        self.root.withdraw()

    def exit_app(self, icon=None, item=None):
        if self.tray_icon: self.tray_icon.stop()
        self.root.destroy()
        os._exit(0)

# --- EVENT BRIDGE ---
class ExcelEventsBridge:
    ui_instance = None
    def OnSheetBeforeDoubleClick(self, Sh, Target, Cancel):
        if self.ui_instance: return self.ui_instance.OnSheetBeforeDoubleClick(Sh, Target, Cancel)
    def OnSheetBeforeRightClick(self, Sh, Target, Cancel):
        if self.ui_instance: return self.ui_instance.OnSheetBeforeRightClick(Sh, Target, Cancel)

def monitor_loop(ui):
    ExcelEventsBridge.ui_instance = ui
    while True:
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            ui.root.after(0, ui.update_connection_status, True)
            
            # Bind events
            # Note: win32com handles reconnections poorly if we don't recreate the bridge
            bridge = win32com.client.WithEvents(excel, ExcelEventsBridge)
            
            print("Connected to Excel.")
            while True:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.1)
                # Check if Excel still exists
                try:
                    _ = excel.Name
                except:
                    break
        except Exception:
            ui.root.after(0, ui.update_connection_status, False)
            time.sleep(3) # Retry every 3 seconds

def main():
    ui = ExcelImageToolUI()
    threading.Thread(target=monitor_loop, args=(ui,), daemon=True).start()
    ui.root.mainloop()

if __name__ == "__main__":
    main()

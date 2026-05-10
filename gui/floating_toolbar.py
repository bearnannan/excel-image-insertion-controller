# =================================================================
# Project: Excel Image Insertion Controller
# Component: Floating Toolbar Widget
# Author: WATCHARA MANADEE
# License: MIT
# =================================================================

import customtkinter as ctk
from tkinterdnd2 import DND_FILES

class FloatingToolbar(ctk.CTkToplevel):
    def __init__(self, master, controller):
        super().__init__(master)
        
        self.controller = controller
        
        self.title("Excel Image Toolbar")
        self.geometry("200x60+100+100")
        self.overrideredirect(True) # Frameless
        self.attributes("-topmost", True)
        
        # Make transparent if possible (Windows specific)
        self.attributes("-alpha", 0.9)
        
        # Make draggable
        self.bind("<ButtonPress-1>", self.start_move)
        self.bind("<B1-Motion>", self.do_move)
        
        # Frame
        self.frame = ctk.CTkFrame(self, corner_radius=15, border_width=2, border_color="#217346")
        self.frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        # Mode Toggle Button
        self.toggle_btn = ctk.CTkButton(
            self.frame, 
            text="Mode: OFF", 
            width=80, 
            height=30,
            command=self.toggle_clicked,
            fg_color="transparent",
            border_width=1,
            text_color=("black", "white")
        )
        self.toggle_btn.pack(side="left", padx=(10, 5), pady=10)
        
        # Drag and Drop Area
        self.dnd_label = ctk.CTkLabel(
            self.frame, 
            text="Drop Img Here", 
            width=80, 
            fg_color=("gray80", "gray20"),
            corner_radius=8
        )
        self.dnd_label.pack(side="right", padx=(5, 10), pady=10, fill="both", expand=True)
        
        # Setup Drag and Drop
        # Note: tkinterdnd2 requires the root to be TkinterDnD.Tk
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self.on_drop)
        except Exception as e:
            print(f"DND not supported: {e}")
            self.dnd_label.configure(text="DND Err")

        self._x = 0
        self._y = 0

    def start_move(self, event):
        self._x = event.x
        self._y = event.y

    def do_move(self, event):
        deltax = event.x - self._x
        deltay = event.y - self._y
        x = self.winfo_x() + deltax
        y = self.winfo_y() + deltay
        self.geometry(f"+{x}+{y}")

    def toggle_clicked(self):
        self.controller.toggle_image_mode()

    def update_mode_status(self, is_on):
        if is_on:
            self.toggle_btn.configure(text="Mode: ON", fg_color="#FFC107", text_color="black")
            self.frame.configure(border_color="#FFC107")
        else:
            self.toggle_btn.configure(text="Mode: OFF", fg_color="transparent", text_color=("black", "white"))
            self.frame.configure(border_color="#217346")

    def on_drop(self, event):
        # The data contains the file path(s). Windows might wrap them in {} if there are spaces.
        files = self.master.tk.splitlist(event.data)
        if files:
            # Pass the first file (or all if batch) to the controller
            # For simplicity, we just insert the first dropped file
            file_path = files[0]
            self.controller.insert_dragged_file(file_path)

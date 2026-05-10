# Excel Image Insertion Controller 🚀

A modern Python GUI application that supercharges Microsoft Excel by adding a robust "Image Insertion Mode". Operating entirely externally, it controls *any* currently open Excel workbook without requiring any VBA macros (`.xlsm`) inside the file itself.

## ✨ Features
- **Zero Excel Modifications**: Works instantly with your currently open `.xlsx` files via COM Automation.
- **4-Click Gesture Control**: Double-click twice rapidly on any cell to toggle the Image Insertion Mode.
- **Floating Toolbar**: A modern, transparent, draggable widget allowing you to drag & drop images directly into the active cell.
- **Batch Insertion**: Select multiple images from the file dialog to insert them in a vertical sequence.
- **Smart Formatting**:
    - Automatic scaling and centering within the cell.
    - Optional "Auto-fit Row Height".
    - Complete `.webp` support (preserves transparency).
- **System Notifications**: Get instant feedback via Windows toast notifications.

---

## 🛠️ Installation & Setup

### Option 1: Run as Standalone Executable (.exe)
1. Double click **`build_exe.bat`** to compile the app.
2. Find the generated `main.exe` in the `dist` folder.
3. Keep the `.exe` anywhere on your system and run it whenever you are using Excel.

### Option 2: Run via Python
Ensure you have Python 3.11+ installed.
1. Install requirements:
   ```bash
   pip install -r requirements.txt
   ```
2. Start the application:
   ```bash
   python main.py
   ```

---

## 📖 How to Use

1. **Start the App**: Open the `main.exe` or run `main.py`. The GUI will open and automatically connect to your open Excel instance.
2. **Toggle Image Mode**:
    - Click **Enable Image Mode** in the GUI or Floating Toolbar.
    - **OR** double-click any Excel cell twice within 1 second.
3. **Insert an Image**:
    - With Image Mode ON, **Right-Click** any Excel cell. A file dialog will open.
    - **OR** Drag & Drop an image file onto the "Drop Img Here" section of the Floating Toolbar.
4. **Batch Insert**: Select multiple images in the file dialog; they will be inserted sequentially downwards from the active cell.

---

## ⚙️ Architecture Requirements
- `customtkinter` (Modern UI)
- `tkinterdnd2` (Drag & Drop Support)
- `pywin32` (Excel COM Automation & Event Hooks)
- `Pillow` (WebP processing)
- `plyer` (Notifications)

---
## 👤 Author
**WATCHARA MANADEE**

# =================================================================
# Project: Excel Image Insertion Controller
# Component: Image Handler (Pillow & Excel Shape Logic)
# Author: WATCHARA MANADEE
# License: MIT
# =================================================================

import os
import tempfile
from PIL import Image

class ImageHandler:
    def __init__(self):
        self.temp_dir = tempfile.gettempdir()
        self.temp_files = []

    def prepare_image(self, file_path):
        """Converts webp to png if necessary and preserves transparency."""
        file_path = os.path.normpath(os.path.abspath(file_path))
        
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return None

        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.webp':
            try:
                img = Image.open(file_path)
                temp_name = f"temp_excel_img_{len(self.temp_files)}_{os.path.getmtime(file_path)}.png"
                temp_path = os.path.join(self.temp_dir, temp_name)
                temp_path = os.path.normpath(os.path.abspath(temp_path))
                
                img.save(temp_path, "PNG")
                if temp_path not in self.temp_files:
                    self.temp_files.append(temp_path)
                return temp_path
            except Exception as e:
                print(f"Error converting webp: {e}")
                return file_path 
        return file_path

    def insert_image(self, sheet, target_range, file_path, settings):
        """
        Deletes old shapes in the range, inserts the new image, and applies settings.
        settings: dict containing:
            - auto_fit_row: bool
            - auto_center: bool
            - padding: int
        """
        try:
            target_addr = target_range.Address
            
            # 1. Cleanup existing images precisely in this cell
            for shape in sheet.Shapes:
                if shape.TopLeftCell.Address == target_addr:
                    shape.Delete()

            # 2. Process image file
            processed_path = self.prepare_image(file_path)
            if not processed_path or not os.path.exists(processed_path):
                print(f"Error: Image file inaccessible: {processed_path}")
                return False

            # 3. Auto fit row height if enabled
            if settings.get('auto_fit_row', False):
                # Simple logic: adjust row height based on image aspect ratio and column width.
                # Since we don't know the image dimensions easily through COM before inserting,
                # we insert first, then adjust.
                pass 

            # Insert original size first
            img = sheet.Shapes.AddPicture(
                Filename=processed_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=target_range.Left,
                Top=target_range.Top,
                Width=-1,
                Height=-1
            )
            
            # Calculate sizing
            padding = settings.get('padding', 2)
            pad_half = padding / 2
            
            # Auto center logic
            auto_center = settings.get('auto_center', True)
            
            # Auto fit row height based on scaled image width
            if settings.get('auto_fit_row', False):
                # Scale image width to column width minus padding
                target_w = target_range.Width - padding
                if target_w > 0:
                    scale_factor = target_w / img.Width
                    new_h = img.Height * scale_factor
                    target_range.RowHeight = new_h + padding

            # Apply final sizing and positioning
            img.LockAspectRatio = False # Allow precise fitting
            
            if auto_center:
                # To center perfectly while maintaining aspect ratio:
                # We fit inside the cell preserving aspect ratio.
                cell_w = target_range.Width - padding
                cell_h = target_range.Height - padding
                
                # Keep original aspect ratio
                img.LockAspectRatio = True
                
                # Set width first, then height limits it
                img.Width = cell_w
                if img.Height > cell_h:
                    img.Height = cell_h
                    
                # Center it
                img.Left = target_range.Left + (target_range.Width - img.Width) / 2
                img.Top = target_range.Top + (target_range.Height - img.Height) / 2
            else:
                # Stretch to fit
                img.Left = target_range.Left + pad_half
                img.Top = target_range.Top + pad_half
                img.Width = target_range.Width - padding
                img.Height = target_range.Height - padding

            # Placement: 1 = xlMoveAndSize
            img.Placement = 1 
            
            return True
            
        except Exception as e:
            print(f"Insertion Error: {e}")
            return False

    def cleanup(self):
        """Removes temporary converted files."""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except:
                pass
        self.temp_files.clear()

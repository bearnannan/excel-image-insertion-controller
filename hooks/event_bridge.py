# =================================================================
# Project: Excel Image Insertion Controller
# Component: Excel COM Event Bridge
# Author: WATCHARA MANADEE
# License: MIT
# =================================================================

import time
import pythoncom

class ExcelEventBridge:
    """
    COM Event handler for Excel.
    This class is instantiated by win32com internally.
    We use class attributes to maintain state and communicate with the main controller.
    """
    controller = None
    last_click_time = 0
    toggle_speed = 1.0

    def OnSheetBeforeDoubleClick(self, Sh, Target, Cancel):
        if not self.controller:
            return

        current_time = time.time()
        
        # Check for 4-click (two double clicks within toggle_speed)
        if current_time - ExcelEventBridge.last_click_time < ExcelEventBridge.toggle_speed and ExcelEventBridge.last_click_time > 0:
            # Trigger toggle
            self.controller.toggle_image_mode()
            ExcelEventBridge.last_click_time = 0
            return True # Cancel the double click action in Excel
        else:
            ExcelEventBridge.last_click_time = current_time
            # If mode is on, we cancel the double click so it doesn't enter edit mode
            if self.controller.is_image_mode_enabled():
                return True
                
        return False

    def OnSheetBeforeRightClick(self, Sh, Target, Cancel):
        if not self.controller:
            return

        if self.controller.is_image_mode_enabled():
            # Trigger insertion logic
            self.controller.handle_image_insertion(Sh, Target)
            return True # Cancel the context menu

        return False

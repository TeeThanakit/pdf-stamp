# main.py

import customtkinter
from tkinter import messagebox
import os
from app_ui import App
from file_processor import resource_path, THAI_FONT_FILE

# ==============================================================================
# START THE APPLICATION
# ==============================================================================
if __name__ == "__main__":
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")

    font_path_to_check = resource_path(os.path.basename(THAI_FONT_FILE))
    
    if not os.path.exists(font_path_to_check):
        messagebox.showwarning("Font Warning", f"ไม่พบไฟล์ฟอนต์ '{os.path.basename(THAI_FONT_FILE)}'\n(โปรแกรมจะรวมไฟล์นี้เข้าไปตอนสร้าง .exe)")
    
    app = App()
    app.mainloop()
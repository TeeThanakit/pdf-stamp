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

    if not os.path.exists(THAI_FONT_FILE):
        messagebox.showerror("ข้อผิดพลาดเกี่ยวกับฟอนต์", f"ไม่พบไฟล์ฟอนต์ '{THAI_FONT_FILE}'\nกรุณาวางไฟล์ในโฟลเดอร์เดียวกับโปรแกรม")
    else:
        app = App()
        app.mainloop()
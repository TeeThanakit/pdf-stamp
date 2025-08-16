# app_ui.py

import customtkinter
import tkinter
from tkinter import filedialog, messagebox
import os
import threading
import fitz
import json

from file_processor import load_product_data, process_pdf_document

SPEC_FILE_CONTENT = """
# -*- mode: python ; coding: utf-8 -*-
import customtkinter

block_cipher = None

added_files = [
    ('Sarabun-Regular.ttf', '.'),
    (customtkinter.__path__[0], 'customtkinter')
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'
)
"""

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("โปรแกรมประทับตรา PDF")
        self.geometry("700x750")
        
        self.main_font = ("Leelawadee UI", 14)
        self.button_font = ("Leelawadee UI", 14, "bold")

        self.excel_path = tkinter.StringVar()
        self.input_files = [] 
        self.output_folder = tkinter.StringVar()
        
        self.config = {
            "whiteout_area": None,
            "default_area": None
        }
        
        self.font_size = 10
        self.whiteout_var = tkinter.BooleanVar()
        
        self.selection_frame = customtkinter.CTkFrame(self)
        self.selection_frame.pack(pady=10, padx=10, fill="x")
        
        customtkinter.CTkLabel(self.selection_frame, text="1. เลือกไฟล์ข้อมูล (Excel/CSV):", font=self.main_font).pack(anchor="w", padx=10)
        customtkinter.CTkEntry(self.selection_frame, textvariable=self.excel_path, width=400, font=self.main_font).pack(side="left", fill="x", expand=True, padx=(10, 5), pady=5)
        customtkinter.CTkButton(self.selection_frame, text="เลือก...", command=self.select_excel_file, font=self.main_font).pack(side="left", padx=(0, 10), pady=5)

        customtkinter.CTkLabel(self, text="2. เลือกไฟล์ PDF:", font=self.main_font).pack(anchor="w", padx=20, pady=(10,0))
        
        self.input_files_textbox = customtkinter.CTkTextbox(self, height=80, font=self.main_font)
        self.input_files_textbox.pack(fill="x", padx=20, pady=5)
        self.input_files_textbox.insert("1.0", "ยังไม่ได้เลือกไฟล์")
        self.input_files_textbox.configure(state="disabled")

        customtkinter.CTkButton(self, text="เลือกไฟล์...", command=self.select_input_files, font=self.main_font).pack(padx=20, pady=5)
        
        customtkinter.CTkLabel(self, text="3. เลือกโฟลเดอร์สำหรับไฟล์ที่เสร็จแล้ว:", font=self.main_font).pack(anchor="w", padx=20, pady=(10,0))
        customtkinter.CTkEntry(self, textvariable=self.output_folder, font=self.main_font).pack(fill="x", padx=20, pady=5)
        customtkinter.CTkButton(self, text="เลือก...", command=self.select_output_folder, font=self.main_font).pack(padx=20, pady=5)
        
        self.settings_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        self.settings_frame.pack(pady=10, padx=20, fill="x")
        self.whiteout_checkbox = customtkinter.CTkCheckBox(
            self.settings_frame, 
            text="ลบข้อมูลเก่า (เทพื้นขาว) และวางตารางไว้ด้านบน",
            variable=self.whiteout_var,
            font=self.main_font
        )
        self.whiteout_checkbox.pack(side="left")
        
        self.run_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        self.run_frame.pack(pady=10, padx=20, fill="x")

        self.run_button = customtkinter.CTkButton(self.run_frame, text="▶️ เริ่มประมวลผล", command=self.start_processing_thread, height=40, font=self.button_font)
        self.run_button.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self.settings_button = customtkinter.CTkButton(self.run_frame, text="ตั้งค่าตำแหน่ง", command=self.open_settings_window, height=40)
        self.settings_button.pack(side="left", expand=True, fill="x", padx=(5, 0))

        self.log_textbox = customtkinter.CTkTextbox(self, state="disabled")
        self.log_textbox.pack(pady=10, padx=20, fill="both", expand=True)

        self.log_textbox.tag_config("error", foreground="#e63946")
        self.log_textbox.tag_config("info", foreground="#0077b6")
        
        self.spec_button = customtkinter.CTkButton(self, text="สร้างไฟล์ .spec (สำหรับทำ .exe)", command=self.create_spec_file)
        self.spec_button.pack(pady=(5, 10), padx=20, fill="x")
        self.load_config()

    def create_spec_file(self):
        try:
            with open("main.spec", "w", encoding="utf-8") as f:
                f.write(SPEC_FILE_CONTENT)
            messagebox.showinfo("สำเร็จ", "สร้างไฟล์ main.spec สำเร็จแล้ว!\nคุณสามารถปิดโปรแกรมนี้ และไปรันคำสั่ง 'pyinstaller main.spec' ใน Terminal ได้เลย")
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถสร้างไฟล์ main.spec ได้: {e}")

    def load_config(self):
        try:
            if os.path.exists("config.json"):
                with open("config.json", "r") as f:
                    self.config = json.load(f)
                    self.log("โหลดตำแหน่งที่บันทึกไว้สำเร็จ")
        except Exception as e:
            self.log(f"ไม่สามารถโหลดไฟล์ config: {e}", tag="error")

    def save_config(self):
        try:
            with open("config.json", "w") as f:
                json.dump(self.config, f, indent=4)
            self.log("บันทึกตำแหน่งใหม่สำเร็จแล้ว", tag="info")
        except Exception as e:
            self.log(f"ไม่สามารถบันทึกไฟล์ config: {e}", tag="error")

    def log(self, message, tag=None):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n", tag)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")
        self.update_idletasks()
    
    def set_buttons_state(self, state):
        self.run_button.configure(state=state)
        self.settings_button.configure(state=state)
        if state == "disabled":
            self.run_button.configure(text="กำลังประมวลผล...")
        else:
            self.run_button.configure(text="▶️ เริ่มประมวลผล")

    def select_excel_file(self):
        path = filedialog.askopenfilename(title="เลือกไฟล์ข้อมูล", filetypes=[("ไฟล์ข้อมูล", "*.xlsx *.xls *.csv"), ("ไฟล์ทั้งหมด", "*.*")])
        if path:
            self.excel_path.set(path)
            self.log("เลือกไฟล์ข้อมูล: " + os.path.basename(path))

    def select_input_files(self):
        paths = filedialog.askopenfilenames(title="เลือกไฟล์ PDF", filetypes=[("ไฟล์ PDF", "*.pdf")])
        if paths:
            self.input_files = list(paths)
            self.log(f"เลือกไฟล์ PDF จำนวน {len(self.input_files)} ไฟล์")
            self.input_files_textbox.configure(state="normal")
            self.input_files_textbox.delete("1.0", "end")
            file_names = [os.path.basename(p) for p in self.input_files]
            display_text = "\n".join(file_names)
            self.input_files_textbox.insert("1.0", display_text)
            self.input_files_textbox.configure(state="disabled")

    def select_output_folder(self):
        path = filedialog.askdirectory(title="เลือกโฟลเดอร์ผลลัพธ์")
        if path:
            self.output_folder.set(path)
            self.log("เลือกโฟลเดอร์ผลลัพธ์: " + path)

    def start_processing_thread(self):
        self.set_buttons_state("disabled")
        processing_thread = threading.Thread(target=self.run_processing_logic)
        processing_thread.daemon = True
        processing_thread.start()
        
    def run_processing_logic(self):
        if not all([self.excel_path.get(), self.input_files, self.output_folder.get()]):
            messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกไฟล์ข้อมูล, ไฟล์ PDF, และโฟลเดอร์ผลลัพธ์")
            self.set_buttons_state("normal")
            return

        should_whiteout = self.whiteout_var.get()
        if should_whiteout and not self.config.get("whiteout_area"):
             messagebox.showwarning("ยังไม่ได้ตั้งค่า", "กรุณากดปุ่ม 'ตั้งค่าตำแหน่ง' เพื่อกำหนด 'พื้นที่สำหรับลบและพิมพ์ทับ' ก่อน")
             self.set_buttons_state("normal")
             return
        if not should_whiteout and not self.config.get("default_area"):
            messagebox.showwarning("ยังไม่ได้ตั้งค่า", "กรุณากดปุ่ม 'ตั้งค่าตำแหน่ง' เพื่อกำหนด 'พื้นที่สำหรับพิมพ์ทับปกติ' ก่อน")
            self.set_buttons_state("normal")
            return

        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        self.log("กำลังเริ่มการทำงาน...")
        
        product_data, message = load_product_data(self.excel_path.get())
        if not product_data:
            self.log(message, tag="error")
            self.set_buttons_state("normal")
            return
        self.log(message)

        pdf_files = self.input_files
        self.log(f"พบไฟล์ PDF {len(pdf_files)} ไฟล์ที่จะประมวลผล...")
        
        work_area_coords = self.config["whiteout_area"] if should_whiteout else self.config["default_area"]
        work_area_rect = fitz.Rect(work_area_coords)

        for pdf_path in pdf_files:
            status_messages = process_pdf_document(
                pdf_path, product_data, self.output_folder.get(), 
                work_area_rect, self.font_size, 
                whiteout=should_whiteout, 
                place_at_top=should_whiteout
            )
            for status in status_messages:
                if status.startswith("ข้าม") or status.startswith("ข้อผิดพลาด") or status.startswith("คำเตือน"):
                    self.log(status, tag="error")
                elif status.startswith("ข้อมูล"):
                    self.log(status, tag="info")
                else:
                    self.log(status)
            
        self.log("--- การประมวลผลเสร็จสมบูรณ์ ---")
        self.set_buttons_state("normal")
        
    def open_settings_window(self):
        sample_pdf = filedialog.askopenfilename(title="เลือกไฟล์ PDF ตัวอย่างเพื่อใช้ตั้งค่า", filetypes=[("PDF Files", "*.pdf")])
        if not sample_pdf:
            return

        for i, area_key in enumerate(["whiteout_area", "default_area"]):
            title_text = "ขั้นตอนที่ 1/2: วาดกรอบ 'สำหรับลบและพิมพ์ทับ'" if i == 0 else "ขั้นตอนที่ 2/2: วาดกรอบ 'สำหรับพิมพ์ทับปกติ'"
            
            rect_coords = self.get_rect_from_popup(sample_pdf, title_text)
            if rect_coords:
                self.config[area_key] = rect_coords
            else:
                self.log("ยกเลิกการตั้งค่าตำแหน่ง", tag="info")
                return
        
        self.save_config()

    def get_rect_from_popup(self, pdf_path, title):
        work_rect_coords = None
        picker_window = customtkinter.CTkToplevel(self)
        picker_window.title(title)
        picker_window.attributes("-topmost", True)

        doc = fitz.open(pdf_path)
        page = doc[0]
        page_original_width = page.rect.width
        
        zoom_factor = min(1800 / page.rect.width, 900 / page.rect.height)
        dpi_value = int(zoom_factor * 72)
        pix = page.get_pixmap(dpi=dpi_value)
        
        doc.close()
        
        canvas = tkinter.Canvas(picker_window, width=pix.width, height=pix.height)
        canvas.image = tkinter.PhotoImage(data=pix.tobytes("ppm"))
        canvas.create_image(0, 0, anchor="nw", image=canvas.image)
        canvas.pack(fill="both", expand=True)

        points = []
        rect_id = None
        
        def on_click(event):
            nonlocal points
            points.append((event.x, event.y))
            if len(points) == 1:
                canvas.create_oval(event.x-3, event.y-3, event.x+3, event.y+3, fill="red", outline="red")
            elif len(points) == 2:
                nonlocal work_rect_coords
                scale_factor = page_original_width / pix.width
                p1_pdf = fitz.Point(points[0][0] * scale_factor, points[0][1] * scale_factor)
                p2_pdf = fitz.Point(points[1][0] * scale_factor, points[1][1] * scale_factor)
                valid_rect = fitz.Rect(p1_pdf, p2_pdf)
                work_rect_coords = [valid_rect.x0, valid_rect.y0, valid_rect.x1, valid_rect.y1]
                picker_window.destroy()

        def on_mouse_move(event):
            nonlocal rect_id
            if len(points) == 1:
                if rect_id:
                    canvas.delete(rect_id)
                rect_id = canvas.create_rectangle(points[0][0], points[0][1], event.x, event.y, outline="red", width=2)

        canvas.bind("<Button-1>", on_click)
        canvas.bind("<Motion>", on_mouse_move)
        
        self.wait_window(picker_window)
        return work_rect_coords
import customtkinter
import tkinter
from tkinter import filedialog, messagebox
import os
import threading
import fitz

from file_processor import load_product_data, process_pdf_document

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("โปรแกรมประทับตรา PDF")
        self.geometry("700x600")
        
        self.main_font = ("Leelawadee UI", 14)
        self.button_font = ("Leelawadee UI", 14, "bold")

        self.excel_path = tkinter.StringVar()
        self.input_files = [] 
        self.output_folder = tkinter.StringVar()
        self.text_position = None
        self.font_size = 12
        
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
        
        self.run_button = customtkinter.CTkButton(self, text="▶️ เริ่มการประมวลผล", command=self.start_processing_thread, height=40, font=self.button_font)
        self.run_button.pack(pady=20, padx=20, fill="x")
        
        self.log_textbox = customtkinter.CTkTextbox(self, state="disabled", height=150, font=self.main_font)
        self.log_textbox.pack(pady=10, padx=20, fill="both", expand=True)

        self.log_textbox.tag_config("error", foreground="#e63946")
        self.log_textbox.tag_config("info", foreground="#0077b6")

    def log(self, message, tag=None):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n", tag)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")
        self.update_idletasks()

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
        self.run_button.configure(state="disabled", text="กำลังประมวลผล...")
        processing_thread = threading.Thread(target=self.run_processing_logic)
        processing_thread.daemon = True
        processing_thread.start()
        
    def run_processing_logic(self):
        if not all([self.excel_path.get(), self.input_files, self.output_folder.get()]):
            messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกไฟล์ข้อมูล, ไฟล์ PDF, และโฟลเดอร์ผลลัพธ์")
            self.run_button.configure(state="normal", text="▶️ เริ่มการประมวลผล")
            return

        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        self.log("กำลังเริ่มการทำงาน...")
        
        product_data, message = load_product_data(self.excel_path.get())
        
        if not product_data:
            self.log(message, tag="error")
            self.run_button.configure(state="normal", text="▶️ เริ่มการประมวลผล")
            return
        self.log(message)

        self.log("กรุณาคลิกบนรูปภาพเพื่อกำหนดตำแหน่งข้อความ")
        self.get_text_position_from_user()
        if not self.text_position:
            self.log("ไม่ได้กำหนดตำแหน่งข้อความ, กำลังยกเลิก", tag="error")
            self.run_button.configure(state="normal", text="▶️ เริ่มการประมวลผล")
            return
        self.log(f"กำหนดตำแหน่งข้อความที่: {self.text_position}")

        pdf_files = self.input_files
        self.log(f"พบไฟล์ PDF {len(pdf_files)} ไฟล์ที่จะประมวลผล...")
        for pdf_path in pdf_files:
            status_messages = process_pdf_document(pdf_path, product_data, self.output_folder.get(), self.text_position, self.font_size)
            for status in status_messages:
                # ### UPDATED LOGIC ###
                # เพิ่มเงื่อนไขสำหรับ "คำเตือน" ให้เป็นสีแดง
                if status.startswith("ข้าม") or status.startswith("ข้อผิดพลาด") or status.startswith("คำเตือน"):
                    self.log(status, tag="error")
                elif status.startswith("ข้อมูล"):
                    self.log(status, tag="info")
                else:
                    self.log(status)
            
        self.log("--- การประมวลผลเสร็จสมบูรณ์ ---")
        self.run_button.configure(state="normal", text="▶️ เริ่มการประมวลผล")

    def get_text_position_from_user(self):
        if not self.input_files:
            return
        first_pdf_path = self.input_files[0]
        
        picker_window = customtkinter.CTkToplevel(self)
        picker_window.title("กำหนดตำแหน่งข้อความ")
        picker_window.attributes("-topmost", True)
        doc = fitz.open(first_pdf_path)
        page = doc[0]
        pix = page.get_pixmap()
        doc.close()
        img = tkinter.PhotoImage(data=pix.tobytes("ppm"))
        canvas = tkinter.Canvas(picker_window, width=pix.width, height=pix.height)
        canvas.create_image(0, 0, anchor="nw", image=img)
        canvas.pack()
        def on_click(event):
            self.text_position = fitz.Point(event.x, event.y)
            picker_window.destroy()
        canvas.bind("<Button-1>", on_click)
        self.wait_window(picker_window)
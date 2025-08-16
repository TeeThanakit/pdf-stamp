import pandas as pd
import fitz
import os
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- Configuration ---
THAI_FONT_FILE = resource_path("Sarabun-Regular.ttf") 
PDF_FONT_NAME = "Sarabun" 

def load_product_data(excel_path):
    """
    Reads the data file (.xlsx, .xls, .csv) and converts it into a Python dictionary.
    """
    try:
        _, file_extension = os.path.splitext(excel_path)
        if file_extension.lower() in ['.xlsx', '.xls']:
            df = pd.read_excel(excel_path, engine='openpyxl')
        elif file_extension.lower() == '.csv':
            df = pd.read_csv(excel_path)
        else:
            return None, f"ประเภทไฟล์ไม่รองรับ: {file_extension}. กรุณาใช้ .xlsx, .xls, หรือ .csv"

        if 'Shopee Order No' not in df.columns or 'ชื่อสินค้า' not in df.columns:
            return None, "ไฟล์ข้อมูลต้องมีคอลัมน์ 'Shopee Order No' และ 'ชื่อสินค้า'"
        
        product_dict = {}
        for _, row in df.iterrows():
            order_no = str(row['Shopee Order No']).strip()
            product_name = str(row['ชื่อสินค้า']).strip()
            if order_no not in product_dict:
                product_dict[order_no] = []
            product_dict[order_no].append(product_name)
        return product_dict, "โหลดไฟล์ข้อมูลสำเร็จ"
    except FileNotFoundError:
        return None, f"ข้อผิดพลาด: ไม่พบไฟล์ที่ {excel_path}"
    except Exception as e:
        return None, f"เกิดข้อผิดพลาดในการอ่านไฟล์ข้อมูล: {e}"

def process_pdf_document(pdf_path, product_data, output_folder, text_position, font_size):
    """
    Processes a multi-page PDF. Skips pages without an Order No.
    Warns about pages where product data is not found.
    """
    doc = None
    output_doc = None
    try:
        doc = fitz.open(pdf_path)
        output_doc = fitz.open()
        statuses = []
        pages_processed = 0

        for page_num, page in enumerate(doc):
            order_no = None
            full_text = page.get_text("text")
            for line in full_text.split('\n'):
                if "Shopee Order No." in line:
                    order_no = line.split("Shopee Order No.")[1].strip()
                    break
            
            if not order_no:
                file_name = os.path.basename(pdf_path)
                statuses.append(f"ข้าม: ไม่พบหมายเลขคำสั่งซื้อในไฟล์ {file_name} หน้าที่ {page_num + 1}")
                continue

            pages_processed += 1
            output_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            new_page = output_doc[-1]

            product_names = product_data.get(order_no)
            
            # ### UPDATED LOGIC ###
            # ถ้าหาข้อมูลสินค้าเจอ ให้เตรียมข้อความเพื่อใส่ลงไป
            if product_names:
                text_to_add = "\n".join(product_names)
                new_page.insert_font(fontfile=THAI_FONT_FILE, fontname=PDF_FONT_NAME)
                new_page.insert_text(
                    text_position,
                    text_to_add,
                    fontname=PDF_FONT_NAME,
                    fontsize=font_size,
                    color=(0, 0, 0)
                )
            else:
                # ถ้าหาข้อมูลสินค้าไม่เจอ ให้สร้างข้อความเตือน และไม่ต้องใส่ข้อความใดๆ ลงบนหน้า PDF
                statuses.append(f"คำเตือน: ไม่พบข้อมูลสินค้าสำหรับหมายเลข {order_no}")

        if pages_processed > 0:
            base_name = os.path.basename(pdf_path)
            file_name, file_ext = os.path.splitext(base_name)
            output_filename = f"{file_name}_processed{file_ext}"
            output_path = os.path.join(output_folder, output_filename)
            output_doc.save(output_path)
            statuses.append(f"สำเร็จ: บันทึกไฟล์ {output_filename} จำนวน {pages_processed} หน้า")
        else:
            statuses.append(f"ข้อมูล: ไม่พบหน้าที่มีหมายเลขคำสั่งซื้อในไฟล์ {os.path.basename(pdf_path)} ไม่ได้สร้างไฟล์ผลลัพธ์")
        
        return statuses

    except Exception as e:
        return [f"ข้อผิดพลาดในการประมวลผลไฟล์ {os.path.basename(pdf_path)}: {e}"]
    
    finally:
        if doc:
            doc.close()
        if output_doc:
            output_doc.close()
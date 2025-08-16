# file_processor.py

import pandas as pd
import fitz
import os
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

THAI_FONT_FILE = resource_path("Sarabun-Regular.ttf")
PDF_FONT_NAME = "Sarabun"

def find_column_names(df_columns):
    name_map = {
        'order_no': ["原始单号", "shopee order no"],
        'product_name': ["品名", "ชื่อสินค้า"],
        'quantity': ["数量", "จำนวน"]
    }
    found_columns = {}
    for key, potential_names in name_map.items():
        for name in potential_names:
            cleaned_df_columns = [str(c).strip().lower() for c in df_columns]
            if name.strip().lower() in cleaned_df_columns:
                original_col_name = df_columns[cleaned_df_columns.index(name.strip().lower())]
                found_columns[key] = original_col_name
                break
    return found_columns

def load_product_data(excel_path):
    try:
        _, file_extension = os.path.splitext(excel_path)
        if file_extension.lower() in ['.xlsx', '.xls']:
            df = pd.read_excel(excel_path, engine='openpyxl')
        elif file_extension.lower() == '.csv':
            df = pd.read_csv(excel_path)
        else:
            return None, f"ประเภทไฟล์ไม่รองรับ: {file_extension}. กรุณาใช้ .xlsx, .xls, หรือ .csv."

        cols = find_column_names(df.columns)
        if len(cols) < 3:
            return None, "ไฟล์ข้อมูลต้องมีคอลัมน์สำหรับ: หมายเลขคำสั่งซื้อ, ชื่อสินค้า, และ จำนวน"
        
        product_dict = {}
        for _, row in df.iterrows():
            order_no = str(row[cols['order_no']]).strip()
            product_name = str(row[cols['product_name']]).strip()
            quantity = str(row[cols['quantity']]).strip()
            if order_no not in product_dict:
                product_dict[order_no] = []
            product_dict[order_no].append([product_name, quantity])
        return product_dict, "โหลดไฟล์ข้อมูลสำเร็จ"
    except Exception as e:
        return None, f"เกิดข้อผิดพลาดในการอ่านไฟล์ข้อมูล: {e}"

def process_pdf_document(pdf_path, product_data, output_folder, work_area_rect, font_size, whiteout, place_at_top):
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

            product_list = product_data.get(order_no)
            
            if product_list:
                if whiteout:
                    new_page.draw_rect(work_area_rect, color=(1, 1, 1), fill=(1, 1, 1))

                headers = ["สินค้า", "จำนวน"]
                col_width_name = work_area_rect.width * 0.75
                qty_start_x = work_area_rect.x0 + col_width_name
                
                # --- UPDATED: ปรับความสูงบรรทัดให้แคบลง ---
                line_height = font_size + 2 
                
                start_point = work_area_rect.tl
                current_y = start_point.y
                
                new_page.insert_font(fontfile=THAI_FONT_FILE, fontname=PDF_FONT_NAME)
                
                header_rect_name = fitz.Rect(work_area_rect.x0, current_y, qty_start_x - 5, current_y + 50)
                header_rect_qty = fitz.Rect(qty_start_x, current_y, work_area_rect.x1, current_y + 50)
                new_page.insert_textbox(header_rect_name, headers[0], fontname=PDF_FONT_NAME, fontsize=font_size, fontfile=THAI_FONT_FILE)
                new_page.insert_textbox(header_rect_qty, headers[1], fontname=PDF_FONT_NAME, fontsize=font_size, fontfile=THAI_FONT_FILE)
                
                current_y += line_height
                header_underline_y = current_y
                new_page.draw_line(
                    fitz.Point(work_area_rect.x0, header_underline_y), 
                    fitz.Point(work_area_rect.x1, header_underline_y)
                )

                current_y += 4
                for name, qty in product_list:
                    name_rect = fitz.Rect(work_area_rect.x0, current_y, qty_start_x - 5, current_y + 50)
                    qty_rect = fitz.Rect(qty_start_x, current_y, work_area_rect.x1, current_y + 50)
                    
                    text_height = new_page.insert_textbox(name_rect, name, fontname=PDF_FONT_NAME, fontsize=font_size, fontfile=THAI_FONT_FILE)
                    new_page.insert_textbox(qty_rect, str(qty), fontname=PDF_FONT_NAME, fontsize=font_size, fontfile=THAI_FONT_FILE)
                    
                    if text_height < (font_size + 1):
                        current_y += line_height
                    else:
                        current_y += (text_height + 4)

            else:
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
        if doc: doc.close()
        if output_doc: output_doc.close()
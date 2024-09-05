import fitz  # ใช้สำหรับการจัดการ PDF รวมถึงการอ่าน แก้ไข และสร้างไฟล์ PDF
from fpdf import FPDF  # ใช้สำหรับการสร้างไฟล์ PDF จากศูนย์ รวมถึงการเพิ่มข้อความ รูปภาพ และองค์ประกอบอื่นๆ
import os  # ใช้สำหรับการทำงานกับระบบปฏิบัติการ

# [1] หาชื่อ field ที่อยู่ใน template pdf
def find_form_fields(pdf_path):
    # จะใช้ fitz ในการอ่าน pdf
    pdf_document = fitz.open(pdf_path)
    form_fields = []
    field_name_counts = {}
    field_name_to_index = {}

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        fields = page.widgets()
        
        if fields:
            for field in fields:
                field_name = field.field_name # ชื่อ field
                field_rect = field.rect # ตำแหน่งของ field ที่อยู่ใน pdf
                x_position = field_rect.x0 # ตำแหน่งของ field ที่อยู่ใน pdf
                y_position = field_rect.y0 # ตำแหน่งของ field ที่อยู่ใน pdf
                width = field_rect.width # ความกว้างขนาดของ field
                height = field_rect.height # ความยาวขนาดของ field

                field_type = field.field_type # ประเภทของ field ดูว่าเป็น Text Field , CheckBox , Image

                if field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                    field_type_str = 'TextField'
                elif field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                    field_type_str = 'Checkbox'
                elif field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
                    field_type_str = 'Radio Button'
                else:
                    field_type_str = 'Unknown'

                # ในกรณีที่ checkbox มีชื่อซ้ำกัน
                if field_name in field_name_counts and field_type_str == 'Checkbox':
                    # จะทำการเปลี่ยน ชื่อ fields checkbox ที่ซ้ำกัน เช่น checkbox มีสองตัวใน pdf
                    field_name_counts[field_name] += 1
                    field_name_with_suffix = f"{field_name}_{field_name_counts[field_name] - 1}"

                    previous_index = field_name_to_index[field_name]
                    form_fields[previous_index] = form_fields[previous_index][:1] + (f"{field_name}_{field_name_counts[field_name]}",) + form_fields[previous_index][2:]
                    unix = 'Yes'
                    # ทำการเปลี่ยนชื่อ field ใน pdf template
                    change_field_name(pdf_path, field_name, previous_index, pdf_path)
                else:
                    field_name_counts[field_name] = 1
                    field_name_with_suffix = field_name
                    unix = 'No'

                field_name_to_index[field_name] = len(form_fields)
                form_fields.append((page_num, field_name_with_suffix, round(x_position, 3), round(y_position, 3), round(width, 3), round(height, 3), field_type_str, unix))

    pdf_document.close()
    return form_fields

# [2] ใช้เปลี่ยนชื่อ field ถ้ามีชื่อที่ซ้ำกัน ในกรณีที่ type ของ field เป็น checkbox
def change_field_name(pdf_path, old_name, new_name, output_path):
    # เปิดไฟล์ pdf template
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        fields = page.widgets()
        
        # ถ้ามีชื่อซ้ำกันจะทำตัวเลขไว้ด้านหลัง เช่น test_1 test_2
        if fields:
            for field in fields:
                if field.field_name == old_name:
                    field.field_name = new_name
                    print(f"Field name changed from {old_name} to {new_name} on page {page_num + 1}")
    
    # save pdf template ที่ชื่อ field เปลี่ยนแล้ว
    if pdf_path == output_path:
        pdf_document.saveIncr()
    else:
        pdf_document.save(output_path)
    
    pdf_document.close()

# [3] สร้างไฟล์เปล่า ที่มี data อยู่ด้านใน
def create_pdf_with_data(form_fields, data_fields, output_path, font_path):
    class PDF(FPDF):
        # เอา margin ของแต่ละหน้าออก ถ้าไม่ได้เอาออกบาง field ที่ map ค่าลงไปจะขึ้นหน้าใหม่หรืออยู่ผิดตำแหน่ง
        def __init__(self):
            super().__init__()
            self.set_margins(0, 0, 0)
            self.set_auto_page_break(auto=False)
            self.current_page = None

        # เพิ่มหน้ากระดาษใหม่เมื่อเริ่มหน้าใหม่ใน PDF
        def add_page_if_needed(self, page_num):
            if self.current_page != page_num:
                self.add_page()
                self.current_page = page_num

        # วาด checkbox
        def draw_checkbox(self, x, y, checkbox_size, boolean_variable):
            self.set_xy(x, y)
            check = "4" if boolean_variable else ""
            self.set_font('ZapfDingbats', '', 10)
            self.cell(checkbox_size, checkbox_size, check, border=0, ln=0, align='C')

    pdf = PDF()

    # กำหนด font ที่จะใช้ map ใน pdf data
    pdf.add_font('THSarabunNew', '', font_path, uni=True)
    pdf.set_font('THSarabunNew', '', 12)

    # วาด cell พร้อมพื้นหลังตามตำแหน่ง ขนาด และประเภทของฟิลด์ที่กำหนด
    def draw_cell_with_background(x, y, width, height, text, bg_color, field_type_str):
        pdf.set_xy(x, y)
        pdf.set_fill_color(*bg_color)
        # ถ้าเป็นรูปภาพ (.png) วางรูปภาพตามตำแหน่งที่กำหนด
        if field_type_str == 'Unknown' and text.endswith(".png"):
            pdf.image(text, x=x, y=y, w=width, h=height)
        # ถ้าเป็นช่องกรอกข้อความ แสดงข้อความใน cell
        elif field_type_str == 'TextField':
            pdf.cell(width, height, text, border=0, ln=0, align='L', fill=False)
        # ถ้าเป็นช่องทำเครื่องหมาย (Checkbox) วาด checkbox ตามค่าที่กำหนด
        elif field_type_str == 'Checkbox':
            checkbox_size = min(width, height)
            is_checked = text.lower() in ['true', 'checked', '1', 'yes']
            pdf.draw_checkbox(x, y, checkbox_size, is_checked)
            pdf.set_font('THSarabunNew', '', 12)
        # กรณีอื่นๆ แสดงข้อความใน cell
        else:
            pdf.cell(width, height, text, border=0, ln=0, align='L', fill=False)

    # กำหนด DPI (จำนวนจุดต่อนิ้ว) และการแปลงจากพิกเซลเป็นมิลลิเมตร
    # โดยที่ต้องแปลงเป็น mm เพราะค่าตำแหน่งที่ได้จาก fitz (PyMuPDF) เป็นพิกเซล
    # แต่ FPDF (ที่ใช้สร้าง PDF) ต้องใช้หน่วยเป็นมิลลิเมตร
    dpi = 72
    px_to_mm = 25.4 / dpi # การแปลงพิกเซลเป็นมิลลิเมตร

    # loop เพื่อประมวลผลฟิลด์ในแต่ละหน้า
    for page_num, field_name, x, y, width, height , field_type_str, unix in form_fields:
        pdf.add_page_if_needed(page_num) # เพิ่มหน้าใหม่หากฟิลด์อยู่ในหน้าที่ต่างจากหน้าปัจจุบัน

        # ถ้าฟิลด์ไม่มีใน data_fields แต่ระบุว่าเป็นฟิลด์แบบมี suffix (unix == 'Yes')
        if field_name not in data_fields and unix == 'Yes':
            data_fields[field_name] = field_name # เพิ่มฟิลด์เข้าไปใน data_fields
            data_fields = process_data_fields(data_fields)

         # ถ้าพบข้อมูลสำหรับฟิลด์นี้ใน data_fields
        if field_name in data_fields:
            value = data_fields[field_name] # ดึงค่าที่จะเติมในฟิลด์

            # แปลงตำแหน่งและขนาดจากพิกเซลเป็นมิลลิเมตร
            x_mm = x * px_to_mm
            y_mm = y * px_to_mm
            width_mm = width * px_to_mm
            height_mm = height * px_to_mm
            background_color = (255, 255, 255) # กำหนดสีพื้นหลังเป็นสีขาว

            # วาด cell พร้อมใส่ข้อมูลตามฟิลด์
            draw_cell_with_background(x_mm, y_mm, width_mm, height_mm, value, background_color, field_type_str)

        else:
            # ถ้าไม่พบข้อมูลสำหรับฟิลด์นี้
            print(f"Warning: No data found for field '{field_name}'")

    pdf.output(output_path)

# [4] ใช้เปลี่ยนค่า checkbox ที่ซ้ำกัน
def process_data_fields(data_fields):
    processed_fields = {} # ใช้เก็บข้อมูลที่ถูกประมวลผล
    base_name_map = {} # ใช้เก็บข้อมูลฟิลด์โดยจัดกลุ่มตามชื่อ base name (ส่วนที่อยู่ก่อน '_')

     # แบ่งชื่อฟิลด์
    for key, value in data_fields.items():
        if '_' in key:
            base_name, suffix = key.rsplit('_', 1)  # แยกฟิลด์เป็น base name และ suffix (เช่น 'field_1', base_name = 'field', suffix = '1')
        else:
            base_name, suffix = key, ''  # ถ้าไม่มี '_', ถือว่า suffix เป็นค่าว่าง
        
        if base_name not in base_name_map:
            base_name_map[base_name] = {}  # ถ้ายังไม่มี base name นี้ใน map ก็สร้างใหม่
        
        base_name_map[base_name][suffix] = value  # เก็บค่าไว้ใน base name map

    for base_name, suffix_map in base_name_map.items():

        if '' in suffix_map:
            processed_fields[base_name] = suffix_map['']
        
        for suffix, value in suffix_map.items():
            if suffix:
                processed_fields[f'{base_name}_{suffix}'] = value

    # ปรับค่าฟิลด์แบบ checkbox ให้สัมพันธ์กัน
    for base_name in list(base_name_map.keys()):
        if base_name in processed_fields:
            value = processed_fields[base_name]
            suffixes = [suffix for suffix in base_name_map.get(base_name, {}).keys() if suffix] # ดึง suffix ทั้งหมดที่ไม่ใช่ค่าว่าง

            # ถ้าฟิลด์หลักมีค่า 'yes' หรือ 'no', จะปรับค่าในฟิลด์ย่อยตาม suffix
            for suffix in suffixes:
                if value == 'yes':
                    if suffix == '1':
                        processed_fields[f'{base_name}_{suffix}'] = 'yes'
                    elif suffix == '2':
                        processed_fields[f'{base_name}_{suffix}'] = 'no'
                elif value == 'no':
                    if suffix == '1':
                        processed_fields[f'{base_name}_{suffix}'] = 'no'
                    elif suffix == '2':
                        processed_fields[f'{base_name}_{suffix}'] = 'yes'

    return processed_fields # คืนค่าข้อมูลที่ประมวลผลแล้ว

# [5] merge pdf data กับ pdf template
def merge_pages(original_pdf_path, data_pdf_path, output_path):
    # เปิดไฟล์ PDF template และ PDF data
    original_pdf = fitz.open(original_pdf_path)
    data_pdf = fitz.open(data_pdf_path)

    # สร้างเอกสาร PDF ใหม่เพื่อเก็บผลลัพธ์ที่ merge กัน
    merged_pdf = fitz.open()

    # ตรวจสอบว่า PDF ทั้งสองมีจำนวนหน้าเท่ากัน
    if len(original_pdf) != len(data_pdf):
        raise ValueError("PDFs must have the same number of pages")

    for page_num in range(len(original_pdf)):
        # โหลดหน้าปัจจุบันจาก PDF ต้นฉบับ
        original_page = original_pdf.load_page(page_num)
        # สร้างหน้าที่ยังใหม่ใน PDF ที่รวมกันด้วยขนาดเดียวกันกับหน้าต้นฉบับ
        merged_page = merged_pdf.new_page(width=original_page.rect.width, height = original_page.rect.height)
        # เพิ่มเนื้อหาจาก PDF ต้นฉบับลงในหน้าที่ยังใหม่
        merged_page.show_pdf_page(merged_page.rect, original_pdf, page_num)
        # เพิ่มเนื้อหาจาก PDF ข้อมูลลงในหน้าที่ยังใหม่เหนือเนื้อหาของ PDF ต้นฉบับ
        merged_page.show_pdf_page(merged_page.rect, data_pdf, page_num)

    # บันทึก PDF ที่รวมกันไปยังที่อยู่เอาต์พุตที่ระบุ
    merged_pdf.save(output_path)
    merged_pdf.close()

    # ตรวจสอบว่าไฟล์ PDF data มีอยู่และลบมันออก
    if os.path.exists(data_pdf_path):
        os.remove(data_pdf_path)
        print(f"{data_pdf_path} has been deleted.")

def main():
    # กำหนด template ที่จะเอามา mapping
    pdf_path = '../PDF template/new_template.pdf'
    output_path = 'output_with_data.pdf'
    merged_output_path = './result/PDFMapper.pdf'
    # กำหนด font ที่จะใช้ map
    font_path = '../fonts/THSarabunNew.ttf'

    # หาชื่อ fields และตำแหน่งของ fields ที่อยู่ใน pdf
    form_fields = find_form_fields(pdf_path)

    # data ที่จะเอาไป map กับ field
    data_fields = {
        'department_manager':'department_manager',
        'insurance_company':'insurance_company',
        'car_code':'car_code',
        'brand_code':'brand_code',
        'code':'code',
        'type_car':'asdas',
        'insureFullName':'insureFullName',
        'amount':'amount',
        'rental_contract':'rental_contract',
        'certificate':'certificate',
        'book':'book',
        'insurFullName':'insurFullName',
        'img_signature':'signature.png',
    }

    # สร้างไฟล์ ที่มีข้อมูลลงไป map แล้ว
    create_pdf_with_data(form_fields, data_fields, output_path, font_path)
    # เอาไฟล์ ที่สร้างขึ้นมาไป merge เข้ากับ template pdf ที่มีอยู่
    merge_pages(pdf_path, output_path, merged_output_path)

if __name__ == "__main__":
    main()

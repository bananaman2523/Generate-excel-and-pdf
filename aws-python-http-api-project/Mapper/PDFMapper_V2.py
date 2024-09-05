import fitz # PyMuPDF สำหรับจัดการไฟล์ PDF

# [1] ฟังก์ชันค้นหาชื่อฟิลด์ใน PDF
def find_form_fields(pdf_path):
    pdf_document = fitz.open(pdf_path) # เปิดไฟล์ PDF
    form_fields = [] # เก็บข้อมูลฟิลด์ฟอร์ม

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        fields = page.widgets()

        if fields:
            for field in fields:
                field_name = field.field_name
                field_rect = field.rect
                
                # ดึงตำแหน่ง (left, top, right, bottom) ของ field
                left = field_rect.x0
                top = field_rect.y0
                right = field_rect.x1
                bottom = field_rect.y1
                field_type = field.field_type

                width = field_rect.width # ความกว้างขนาดของ field
                height = field_rect.height # ความยาวขนาดของ field

                # ตรวจสอบประเภทของ fields
                if field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                    field_type_str = 'TextField'
                elif field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                    field_type_str = 'Checkbox'
                elif field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
                    field_type_str = 'Radio Button'
                else:
                    field_type_str = 'Unknown'
                
                # หากเป็นฟิลด์ที่ไม่รู้จักถือว่าเป็น field image
                if field_type_str == 'Unknown':
                    field_type_str = 'ImageField'

                # เพิ่มฟิลด์ลงในรายการ
                form_fields.append((page_num, field_name, field_type_str, round(left, 3),round(top, 3),round(right, 3),round(bottom, 3),round(width, 3),round(height, 3)))

    pdf_document.close()
    return form_fields # ส่งคืนข้อมูลฟิลด์

# [2] map data ลง file pdf
def create_pdf_with_data(form_fields, data_fields, pdf_path, output_path, font_path):
    doc = fitz.open(pdf_path) # เปิดไฟล์ PDF

    font_file1 = font_path # กำหนดไฟล์ฟอนต์ที่ต้องการใช้
    ZapfDingbats = '../fonts/ZapfDingbats.ttf'
    for page_num, field_name, field_type_str, left, top, right, bottom, width, height in form_fields:
        page = doc.load_page(page_num)
        if field_name in data_fields:
            page.insert_font(fontfile=font_file1, fontname="F0")
            page.insert_font(fontfile=ZapfDingbats, fontname="checkbox")
            value = data_fields[field_name]
            if field_type_str == 'ImageField':
                rect = fitz.Rect(left, top, right, bottom)
                page.insert_image(
                    rect,
                    filename=value
                )
            elif field_type_str == 'Checkbox':
                # ใช้ insert_htmlbox ในการเขียน data ลงไปใน pdf
                # rect กำหนดตำแหน่งที่เราจะเขียน data
                rect = fitz.Rect(left, top, right, bottom)
                is_checked = value.lower() in ['true', 'checked', '1', 'yes']
                check = "4" if is_checked else ""
                html_content = f"""
                <html>
                <head>
                    <style>
                        @font-face {{
                            font-family: 'ZapfDingbats';
                            src: url('../fonts/ZapfDingbats.ttf') format('truetype');
                        }}

                        body {{
                            font-family: 'ZapfDingbats', sans-serif;
                            text-align: center;
                        }}
                    </style>
                </head>
                <body>
                    <div class="wrap-word">
                        {check}
                    </div>
                </body>
                </html>
                """
                page.insert_htmlbox(rect, html_content)
            else:
                # ใช้ insert_htmlbox ในการเขียน data ลงไปใน pdf
                # rect กำหนดตำแหน่งที่เราจะเขียน data
                rect = fitz.Rect(left, top, right, bottom)
                html_content = f"""
                <html>
                <head>
                    <style>
                        .wrap-word {{
                            width: {width};
                            overflow-wrap: break-word;

                        }}
                    </style>
                </head>
                <body>
                    <div class="wrap-word">
                        {value}
                    </div>
                </body>
                </html>
                """
                # fitz สามารถใช้ css กำหนด style ของแต่ละ field ได้
                page.insert_htmlbox(rect, html_content)
        else:
            print(f"'{field_name}' : '{field_name}',")  # แจ้งเตือนถ้าไม่มีข้อมูลในฟิลด์


    doc.save(output_path) # บันทึกไฟล์ PDF ที่สร้างใหม่
    doc.close() # ปิดไฟล์ PDF

# [3] ลบ field ออกจาก template pdf
def remove_fields_from_pdf(input_pdf, output_pdf):
    pdf_document = fitz.open(input_pdf)

    for page_number in range(len(pdf_document)):
        page = pdf_document.load_page(page_number)

        for field in page.widgets():
            page.delete_widget(field)
    if input_pdf == output_pdf:
        pdf_document.saveIncr()
    else:
        pdf_document.save(output_pdf, deflate=True)
        pdf_document.close()

def main():
    # กำหนด template ที่จะเอามา mapping
    pdf_path = '../PDF template/test_pdf.pdf'
    output_path = './result/PDFMapper_V2.pdf'
    # กำหนด font ที่จะใช้ map
    font_path = '../fonts/THSarabunNew.ttf'

    # หาชื่อ fields และตำแหน่งของ fields ที่อยู่ใน pdf
    form_fields = find_form_fields(pdf_path)

    # data ที่จะเอาไป map กับ field
    data_fields = {
        'insured_full_name' : "insured_full_name",
        'insured_address' : 'insured_address',
        'insured_birthdate' : 'insured_birthdate',
        'insured_country' : 'insured_country',
        'insured_city' : 'insured_city',
        'insured_tax_residence_in_countires_yes' : 'Yes',
        'insured_tax_residence_in_countires_no' : 'No',
        'insured_country_tax_residence_1' : 'insured_country_tax_residence_1',
        'insured_tin_number_1' : 'insured_tin_number_1',
        'insured_reason_1' : 'insured_reason_1',
        'insured_explain_reason_1' : 'insured_explain_reason_1',
        'insured_country_tax_residence_2' : 'insured_country_tax_residence_2',
        'insured_tin_number_2' : 'insured_tin_number_2',
        'insured_reason_2' : 'insured_reason_2',
        'insured_explain_reason_2' : 'insured_explain_reason_2',
        'insured_country_tax_residence_3' : 'insured_country_tax_residence_3',
        'insured_tin_number_3' : 'insured_tin_number_3',
        'insured_reason_3' : 'insured_reason_3',
        'insured_explain_reason_3' : 'insured_explain_reason_3',
        'insured_country_tax_residence_4' : 'insured_country_tax_residence_4',
        'insured_tin_number_4' : 'insured_tin_number_4',
        'insured_reason_4' : 'insured_reason_4',
        'insured_explain_reason_4' : 'insured_explain_reason_4',
        'insured_country_tax_residence_5' : 'insured_country_tax_residence_5',
        'insured_tin_number_5' : 'insured_tin_number_5',
        'insured_reason_5' : 'insured_reason_5',
        'insured_explain_reason_5' : 'insured_explain_reason_5',
        'insured_date' : 'insured_date',
        'payor_date' : 'payor_date',
        'img_payor_signature' : 'signature.png',
        'img_insured_signature' : 'signature.png',
    }

    # สร้างไฟล์ ที่มีข้อมูลลงไป map แล้ว
    create_pdf_with_data(form_fields, data_fields, pdf_path, output_path, font_path) # สร้าง PDF พร้อมใส่ข้อมูล
    remove_field = "./result/PDFMapper_V2.pdf" # ไฟล์ PDF ที่ต้องลบ field
    remove_fields_from_pdf(output_path, remove_field) # ลบฟิลด์ออกจาก PDF


if __name__ == "__main__":
    main()

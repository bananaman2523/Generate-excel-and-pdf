import pandas as pd # dataframe ที่จะเอาไปรวมกับ template
import os
from fpdf import FPDF # ใช้ gen ไฟล์ PDF
from exports.template import template_pdf # เอา template ของ PDF

def TABLE_dataframe(dataframe, template):
    """
    แปลง DataFrame เป็น tuple ของแถวที่จัดรูปแบบสำหรับ PDF รวมถึงส่วนหัว

    Args:
    dataframe (pd.DataFrame): DataFrame ที่จะแปลง
    template (dict): เทมเพลตที่มีสไตล์คอลัมน์และตัวเลือกการจัดรูปแบบอื่นๆ

    return:
    tuple: โดยที่ tuple ด้านในแต่ละตัวแสดงถึงแถวของข้อมูล
    """
    columns = []
    # แยกชื่อคอลัมน์จาก template
    for i, data in enumerate(template['columns_styles']):
        columns.append(data['columns'])
    header = tuple(columns)

    TABLE_DATA = [header]
    # แปลงแถว DataFrame แต่ละแถวให้เป็นทูเพิลของสตริง
    for _, row in dataframe.iterrows():
        TABLE_DATA.append(tuple(row.astype(str)))

    TABLE_DATA = tuple(TABLE_DATA)

    return TABLE_DATA

def multi_cell_row(pdf, cell_widths, height, data, styles=None, align_center=False):
    """
    วาดแถวของเซลล์ใน PDF จัดการข้อความหลายบรรทัดและใช้รูปแบบ

    อาร์กิวเมนต์:
    pdf (FPDF): อ็อบเจ็กต์ PDF ที่จะวาด
    cell_widths (list): รายการความกว้างของเซลล์สำหรับแถว
    height (int): ความสูงของเซลล์
    data (list): รายการข้อมูลที่จะแสดงในเซลล์
    styles (list, ทางเลือก): รายการพจนานุกรมรูปแบบสำหรับแต่ละเซลล์ ค่าเริ่มต้นคือ None
    align_center (bool, ทางเลือก): กำหนดว่าจะให้แถวอยู่กึ่งกลางบนหน้าหรือไม่ ค่าเริ่มต้นคือ False

    return:
    float: พิกัด y หลังจากวาดแถว
    """
    x = pdf.get_x()
    y = pdf.get_y()
    max_height = 0

    if styles is None:
        styles = [{}] * len(cell_widths)

    total_width = sum(cell_widths)
    if align_center:
        x = (pdf.w - total_width) / 2
        pdf.set_x(x)

    # วาดแต่ละ cell
    for i, (width, style) in enumerate(zip(cell_widths, styles)):
        pdf.set_font(style.get('font', 'THSarabunNew'), style.get('style', ''), style.get('size', 10))
        # กำหนดสี
        color = style.get('color', (0, 0, 0))
        pdf.set_text_color(*color)
        # กำหนดพื้นหลัง
        bg_color = style.get('bg_color', (255, 255, 255))
        pdf.set_fill_color(*bg_color)
        pdf.multi_cell(width, height, data[i], border=0, align=style.get('align', "C"), fill=True)
        if pdf.get_y() - y > max_height:
            max_height = pdf.get_y() - y
        pdf.set_xy(x + sum(cell_widths[:i+1]), y)
    # วาด border ของ cell
    for i in range(len(cell_widths) + 1):
        pdf.line(x + sum(cell_widths[:i]), y, x + sum(cell_widths[:i]), y + max_height)
    pdf.line(x, y, x + total_width, y)
    pdf.line(x, y + max_height, x + total_width, y + max_height)

    pdf.set_y(y + max_height)
    return pdf.get_y()

class PDF(FPDF):
    def __init__(self, template, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.template = template
    # กำหนด header ของเอกสาร
    def header(self):
        for i, data in enumerate(self.template['header_template']):
            self.set_font('THSarabunNew', 'B', size=data['style'].get('font_size', 16))
            self.cell(0, 10, data['header'], 0, 1, 'C')
    # กำหนด footer ของเอกสาร
    def footer(self):
        self.set_y(-15)
        self.set_font('THSarabunNew', 'B', size=10)
        self.set_text_color(0, 0, 0)
        
        left_text = "User : Test Test"
        center_text = f'หน้าที่ {self.page_no()} / {{nb}}'
        right_text = "Create At: 28/08/2567"
        
        left_width = 50
        center_width = 100
        right_width = 50

        pdf_width = self.w - 2 * self.l_margin
        left_x = self.l_margin
        center_x = left_x + left_width
        right_x = pdf_width - right_width + self.l_margin
        
        self.set_xy(left_x, self.h - 15)
        self.cell(left_width, 10, left_text, 0, 0, 'L')
        
        self.set_xy(center_x, self.h - 15)
        self.cell(center_width, 10, center_text, 0, 0, 'C')
        
        self.set_xy(right_x, self.h - 15)
        self.cell(right_width, 10, right_text, 0, 0, 'R')

def generate_pdf_from_dataframe(dataframe, file_name):
    """
    สร้าง PDF จาก DataFrame โดยใช้ template ที่ระบุ

    Args:
    dataframe (pd.DataFrame): DataFrame ที่จะแปลงเป็นไฟล์ PDF
    file_name (str): ชื่อฐานของไฟล์ PDF เอาต์พุต
    """
    dataframe = pd.DataFrame(dataframe)
    template = template_pdf(file_name)

    pdf = PDF(template)

    pdf_file_name = file_name + '.pdf'
    
    base_fonts_path = './fonts'
    pdf.add_font('THSarabunNew', '', os.path.join(base_fonts_path, 'THSarabunNew.ttf'))
    pdf.add_font('THSarabunNew', 'B', os.path.join(base_fonts_path, 'THSarabunNew Bold.ttf'))

    pdf.alias_nb_pages()
    if file_name == 'RPCL002':
        pdf.add_page(orientation='L')
    else:
        pdf.add_page(orientation='P')

    TABLE_DATA = TABLE_dataframe(dataframe, template)
    cell_widths = template.get('cell_widths', [9, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18])
    row_height = template.get('row_height', 5)

    header_styles = [
        {'font': 'THSarabunNew', 'style': 'B', 'size': 10, 'align': 'C'} for _ in cell_widths
    ]
    row_styles = [
        {'font': 'THSarabunNew', 'style': '', 'size': 10, 'align': 'L'} for _ in cell_widths
    ]

    # อัปเดตสไตล์ส่วนหัวและแถวตามเทมเพลต
    for col_style in template['columns_styles']:
        col_name = col_style['columns']
        style = col_style['style']
        if col_name in TABLE_DATA[0]:
            col_index = TABLE_DATA[0].index(col_name)
            header_styles[col_index].update(style)

    for row_style in template['rows_styles']:
        row_name = row_style['rows']
        style = row_style['style']
        if row_name in TABLE_DATA[0]:
            col_index = TABLE_DATA[0].index(row_name)
            row_styles[col_index].update(style)

    # แมปชื่อคอลัมน์กับส่วนหัว
    mapped_headers = []
    for col_name in TABLE_DATA[0]:
        mapped_headers.append(template['header_mapping'].get(col_name, col_name))

    pdf.set_font("THSarabunNew", 'B', size=10)
    # วาด header row
    multi_cell_row(pdf, cell_widths, height=row_height, data=mapped_headers, styles=header_styles)

    pdf.set_font("THSarabunNew", size=10)
    # วาดข้อมูลแต่ละแถวโดยจัดการแบ่งหน้าตามความจำเป็น
    for row in TABLE_DATA[1:]:
        if pdf.get_y() + row_height > 270:
            pdf.add_page()
            pdf.set_font("THSarabunNew", 'B', size=10)
            multi_cell_row(pdf, cell_widths, height=row_height, data=mapped_headers, styles=header_styles)
            pdf.set_font("THSarabunNew", size=10)
        multi_cell_row(pdf, cell_widths, height=row_height, data=row, styles=row_styles)

    # save pdf file
    pdf.output(pdf_file_name)
import pandas as pd
import jpype
import os
from fpdf import FPDF
from exports.template import template_pdf

def TABLE_dataframe(dataframe, template):
    columns = []
    for i, data in enumerate(template['columns_styles']):
        columns.append(data['columns'])
    header = tuple(columns)

    TABLE_DATA = [header]

    for _, row in dataframe.iterrows():
        TABLE_DATA.append(tuple(row.astype(str)))

    TABLE_DATA = tuple(TABLE_DATA)

    return TABLE_DATA

def multi_cell_row(pdf, cell_widths, height, data, styles=None, align_center=False):
    x = pdf.get_x()
    y = pdf.get_y()
    max_height = 0

    if styles is None:
        styles = [{}] * len(cell_widths)

    total_width = sum(cell_widths)
    if align_center:
        x = (pdf.w - total_width) / 2
        pdf.set_x(x)

    for i, (width, style) in enumerate(zip(cell_widths, styles)):
        pdf.set_font(style.get('font', 'THSarabunNew'), style.get('style', ''), style.get('size', 10))
        
        color = style.get('color', (0, 0, 0))
        pdf.set_text_color(*color)

        bg_color = style.get('bg_color', (255, 255, 255))
        pdf.set_fill_color(*bg_color)
        pdf.multi_cell(width, height, data[i], border=0, align=style.get('align', "C"), fill=True)
        if pdf.get_y() - y > max_height:
            max_height = pdf.get_y() - y
        pdf.set_xy(x + sum(cell_widths[:i+1]), y)

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

    def header(self):
        for i, data in enumerate(self.template['header_template']):
            self.set_font('THSarabunNew', 'B', size=data['style'].get('font_size', 16))
            self.cell(0, 10, data['header'], 0, 1, 'C')

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

    mapped_headers = []
    for col_name in TABLE_DATA[0]:
        mapped_headers.append(template['header_mapping'].get(col_name, col_name))

    pdf.set_font("THSarabunNew", 'B', size=10)
    multi_cell_row(pdf, cell_widths, height=row_height, data=mapped_headers, styles=header_styles)

    pdf.set_font("THSarabunNew", size=10)
    for row in TABLE_DATA[1:]:
        if pdf.get_y() + row_height > 270:
            pdf.add_page()
            pdf.set_font("THSarabunNew", 'B', size=10)
            multi_cell_row(pdf, cell_widths, height=row_height, data=mapped_headers, styles=header_styles)
            pdf.set_font("THSarabunNew", size=10)
        multi_cell_row(pdf, cell_widths, height=row_height, data=row, styles=row_styles)

    pdf.output(pdf_file_name)
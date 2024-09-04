import fitz
from fpdf import FPDF
import os

def change_field_name(pdf_path, old_name, new_name, output_path):
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        fields = page.widgets()
        
        if fields:
            for field in fields:
                if field.field_name == old_name:
                    field.field_name = new_name
                    print(f"Field name changed from {old_name} to {new_name} on page {page_num + 1}")
    
    if pdf_path == output_path:
        pdf_document.saveIncr()
    else:
        pdf_document.save(output_path)
    
    pdf_document.close()

def find_form_fields(pdf_path):
    pdf_document = fitz.open(pdf_path)
    form_fields = []
    field_name_counts = {}
    field_name_to_index = {}

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        fields = page.widgets()
        
        if fields:
            for field in fields:
                field_name = field.field_name
                field_rect = field.rect
                x_position = field_rect.x0
                y_position = field_rect.y0
                width = field_rect.width
                height = field_rect.height

                field_type = field.field_type

                if field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                    field_type_str = 'TextField'
                elif field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                    field_type_str = 'Checkbox'
                elif field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
                    field_type_str = 'Radio Button'
                # elif field_type == fitz.PDF_WIDGET_TYPE_BUTTON:
                #     field_type_str = 'Button'
                else:
                    field_type_str = 'Unknown'
                
                if field_name in field_name_counts:
                    field_name_counts[field_name] += 1
                    field_name_with_suffix = f"{field_name}_{field_name_counts[field_name] - 1}"

                    previous_index = field_name_to_index[field_name]
                    form_fields[previous_index] = form_fields[previous_index][:1] + (f"{field_name}_{field_name_counts[field_name]}",) + form_fields[previous_index][2:]
                    change_field_name(pdf_path, field_name, previous_index, pdf_path)
                else:
                    field_name_counts[field_name] = 1
                    field_name_with_suffix = field_name

                field_name_to_index[field_name] = len(form_fields)
                form_fields.append((page_num, field_name_with_suffix, round(x_position, 3), round(y_position, 3), round(width, 3), round(height, 3), field_type_str))

    pdf_document.close()
    return form_fields

def process_data_fields(data_fields):
    processed_fields = {}    
    base_name_map = {}

    for key, value in data_fields.items():
        if '_' in key:
            base_name, suffix = key.rsplit('_', 1)
        else:
            base_name, suffix = key, ''
        
        if base_name not in base_name_map:
            base_name_map[base_name] = {}
        
        base_name_map[base_name][suffix] = value

    for base_name, suffix_map in base_name_map.items():

        if '' in suffix_map:
            processed_fields[base_name] = suffix_map['']
        
        for suffix, value in suffix_map.items():
            if suffix:
                processed_fields[f'{base_name}_{suffix}'] = value

    for base_name in list(base_name_map.keys()):
        if base_name in processed_fields:
            value = processed_fields[base_name]
            suffixes = [suffix for suffix in base_name_map.get(base_name, {}).keys() if suffix]

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

    return processed_fields

def create_pdf_with_data(form_fields, data_fields, output_path, font_path):
    class PDF(FPDF):
        def __init__(self):
            super().__init__()
            self.set_margins(0, 0, 0)
            self.set_auto_page_break(auto=False)
            self.current_page = None

        def add_page_if_needed(self, page_num):
            if self.current_page != page_num:
                self.add_page()
                self.current_page = page_num

        def draw_checkbox(self, x, y, checkbox_size, boolean_variable):
            self.set_xy(x, y)
            check = "4" if boolean_variable else ""
            self.set_font('ZapfDingbats', '', 10)
            self.cell(checkbox_size, checkbox_size, check, border=0, ln=0, align='C')

    pdf = PDF()

    pdf.add_font('THSarabunNew', '', font_path, uni=True)
    pdf.set_font('THSarabunNew', '', 12)

    def draw_cell_with_background(x, y, width, height, text, bg_color, field_type_str):
        pdf.set_xy(x, y)
        pdf.set_fill_color(*bg_color)
        if field_type_str == 'Unknown' and text.endswith(".png"):
            pdf.image(text, x=x, y=y, w=width, h=height)
        elif field_type_str == 'TextField':
            pdf.cell(width, height, text, border=0, ln=0, align='L', fill=False)
        elif field_type_str == 'Checkbox':
            checkbox_size = min(width, height)
            is_checked = text.lower() in ['true', 'checked', '1', 'yes']
            pdf.draw_checkbox(x, y, checkbox_size, is_checked)
            pdf.set_font('THSarabunNew', '', 12)
        else:
            pdf.cell(width, height, text, border=0, ln=0, align='L', fill=False)

    dpi = 72
    px_to_mm = 25.4 / dpi

    for page_num, field_name, x, y, width, height , field_type_str in form_fields:
        pdf.add_page_if_needed(page_num)
        if field_name not in data_fields:
            data_fields[field_name] = field_name
            data_fields = process_data_fields(data_fields)

        if field_name in data_fields:
            value = data_fields[field_name]
            x_mm = x * px_to_mm
            y_mm = y * px_to_mm
            width_mm = width * px_to_mm
            height_mm = height * px_to_mm
            background_color = (255, 255, 255)

            draw_cell_with_background(x_mm, y_mm, width_mm, height_mm, value, background_color, field_type_str)

        else:
            # print(f"'{field_name}' : '{field_name}',")
            print(f"Warning: No data found for field '{field_name}'")

    pdf.output(output_path)

def merge_pages(original_pdf_path, data_pdf_path, output_path):
    original_pdf = fitz.open(original_pdf_path)
    data_pdf = fitz.open(data_pdf_path)

    merged_pdf = fitz.open()

    if len(original_pdf) != len(data_pdf):
        raise ValueError("PDFs must have the same number of pages")

    for page_num in range(len(original_pdf)):
        original_page = original_pdf.load_page(page_num)
        data_page = data_pdf.load_page(page_num)

        merged_page = merged_pdf.new_page(width=original_page.rect.width, height = original_page.rect.height)
        
        merged_page.show_pdf_page(merged_page.rect, original_pdf, page_num)
        
        merged_page.show_pdf_page(merged_page.rect, data_pdf, page_num)

    merged_pdf.save(output_path)
    merged_pdf.close()

    if os.path.exists(data_pdf_path):
        os.remove(data_pdf_path)
        print(f"{data_pdf_path} has been deleted.")

pdf_path = 'vulnerable.pdf'
output_path = 'output_with_data.pdf'
merged_output_path = 'merged_output.pdf'
font_path = '../fonts/THSarabunNew.ttf'

form_fields = find_form_fields(pdf_path)

data_fields = {
    'applicationNumber' : 'applicationNumber',
    'insuredFullName' : 'insuredFullName',
    'ageSixtyAndAbove' : 'yes',
    'noInvestmentExperience' : 'yes',
    'individualsWithDisabilities' : 'yes',
    'vulnerableGroup' : 'yes',
    'trustedIndividualFullName' : 'trustedIndividualFullName',
    'relationshipWithInsured' : 'relationshipWithInsured',
    'phoneNumber' : 'phoneNumber',
    'insuredFullNameSignature' : 'insuredFullNameSignature',
    'img_PayorSignature' : 'signature.png',
    'representativeSignature' : 'representativeSignature',
    'img_InsuredSignature' : 'signature.png',
    'img_ConfidantSignature' : 'signature.png',
    'trustedIndividualFullNameSignature' : 'trustedIndividualFullNameSignature',
    'img_AgentSignature' : 'signature.png',
    'agentSignature' : 'agentSignature',
    'confirmingTheConditions' : 'no',
}

create_pdf_with_data(form_fields, data_fields, output_path, font_path)
merge_pages(pdf_path, output_path, merged_output_path)
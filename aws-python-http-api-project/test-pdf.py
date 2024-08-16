import base64
import os
import subprocess

def xlsx_to_base64(file_path):
    with open(file_path, "rb") as file:
        encoded_string = base64.b64encode(file.read()).decode("utf-8")
    return encoded_string

def base64_to_xlsx(base64_string, output_path):
    decoded_data = base64.b64decode(base64_string)
    with open(output_path, "wb") as file:
        file.write(decoded_data)

def xlsx_to_pdf(xlsx_path, pdf_path):
    libreoffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    subprocess.run([libreoffice_path, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(pdf_path), xlsx_path])

# Example usage:
xlsx_file_path = "RPCL001.xlsx"
base64_string = xlsx_to_base64(xlsx_file_path)
print("Base64 encoded string:", base64_string)

xlsx_output_path = "output.xlsx"
pdf_output_path = "output.pdf"

base64_to_xlsx(base64_string, xlsx_output_path)
xlsx_to_pdf(xlsx_output_path, pdf_output_path)

print(f"PDF file created at {pdf_output_path}")

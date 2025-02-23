
from pathlib import Path
import traceback
import pymupdf
import re
import csv

# Get the current directory
source_directory = Path(r"[Your file path]")
target_csv_file = "001_costs.csv"
files_to_ignore = {target_csv_file, "Extract_Cost_From_PDF.py"}

def main(source_directory, target_csv_file, files_to_ignore):
    # List all files in the source_directory 
    all_pdf_file_names = [f for f in source_directory.iterdir() if f.is_file() and f.name not in files_to_ignore and f.name.endswith(".pdf")]
    
    # Create Headers
    headers_row = ["Filename", "BL", "PDF Invoice Type","invoice_type_from_title", "PDF Total", "total_from_title"]

    rows = []
    
    print("Files in the source_directory:")
    for file_name in all_pdf_file_names:
        print(file_name.name)
        
        # Find BL number
        parts = file_name.name.replace(".pdf", "").split("+")
        bl = parts[0]
        if "_" in bl:
            bl = parts[0].split("_")[1]
        invoice_type_from_title = parts[1]
        total_from_title = parts[2]
      
         
         
        # Read pdf text
        total = "unknown"
        invoice_type = "unknown"
        document = pymupdf.open(file_name)
        for page in document:
            text = page.get_text()
            print(text)
            
            # Invoice Type
            demurrage = re.compile(r"DEMURRAGE INVOICE", re.MULTILINE)
            if demurrage.search(text):
                invoice_type = "DEMURRAGE INVOICE"
                
            detention = re.compile(r"DETENTION INVOICE", re.MULTILINE)
            if detention.search(text):
                invoice_type = "DETENTION INVOICE"
                
            import_invoice = re.compile(r"IMPORT INVOICE|Import Invoice", re.MULTILINE)
            if import_invoice.search(text):
                invoice_type = "IMPORT INVOICE"
            
            eglv = re.compile(r"TOTAA?L\nEUR\n([\d,. ]+)", re.MULTILINE)
            total_eglv = eglv.search(text)
            if total_eglv:
                total = total_eglv.groups()[0]
                print(total)
                break
            orionco = re.compile(r"Totaal:\n.*\nTOGRO\nEUR\n([\d,. ]+)", re.MULTILINE)
            total_orionco = orionco.search(text)
            if total_orionco:
                total = total_orionco.groups()[0]
                print(total)
                break
            maersk = re.compile(r"Total Payable Amount\nEUR\n([\d,. ]+)", re.MULTILINE)
            total_maersk = maersk.search(text)
            if total_maersk:
                total = total_maersk.groups()[0]
                print(total)
                break
            bolore = re.compile(r"TO BE PAID                EUR    +([\d,. ]+)", re.MULTILINE)
            total_bolore = bolore.search(text)
            if total_bolore:
                total = total_bolore.groups()[0]
                print(total)
                break
            exofex = re.compile(r"Total EUR\n([\d,. ]+)", re.MULTILINE)
            total_exofex = exofex.search(text)
            if total_exofex:
                total = total_exofex.groups()[0]
                print(total)
                break
            msc = re.compile(r"Total incl. VAT\n+([\d,. ]+)", re.MULTILINE)
            total_msc = msc.search(text)
            if total_msc:
                total = total_msc.groups()[0]
                print(total)
                break
            

        row = [file_name.name, bl, invoice_type, invoice_type_from_title, total, total_from_title]
        print(row)
        # raise Exception("stop")


        rows.append(row)
        
        
    with open(source_directory / target_csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, dialect='excel', quotechar='"', quoting=csv.QUOTE_STRINGS,)
        writer.writerow(headers_row)
        for row in rows:
            print(row)
            writer.writerow(row)

    

if __name__ == "__main__":
    try:
        main(source_directory, target_csv_file, files_to_ignore)
    except e:
        print("Error!", e, traceback.format_exc())
        input("Error: fix it and try again after closing this terminal")
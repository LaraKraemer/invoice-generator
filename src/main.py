import pandas as pd
import glob
from fpdf import FPDF   
from pathlib import Path
    

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Print the filepath being processed
    print("Processing file:", filepath)
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    # create pdf file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    # extract filename and invoice number 
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    invoice_date = filename.split("-")[1]
    # add content to pdf
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1) 
    
    # add content to pdf
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {invoice_date}", ln=2)
    
    
    pdf.output(f"PDFs/{filename}.pdf")


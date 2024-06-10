import pandas as pd
import glob
from fpdf import FPDF   
from pathlib import Path
    

filepaths = glob.glob("invoices/*.xlsx")


for filepath in filepaths:
    # Print the filepath being processed
    print("Processing file:", filepath)
    
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
    pdf.cell(w=50, h=10, txt=f"Date {invoice_date}", ln=2)
    
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    # add header 
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(100, 80, 80)
    
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)
    
    
    # add rows to the table 
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8,)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)   

    # add total sum 
    sum_total_price = int(df["total_price"].sum())
    print(sum_total_price)
    pdf.set_font(family="Times", size=8, style="B")
    pdf.set_text_color(100, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(sum_total_price), border=1, ln=1)  
    
    # add description 
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(100, 80, 80)
    pdf.cell(w=70, h=20, txt=f"The total price is {sum_total_price} EUR, please transfer until the end of the month.", ln=1) 
    
    # add company name and logo
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(100, 80, 80)
    pdf.cell(w=25, h=20, txt="PythonHow") 
    pdf.image("pythonhow.png", w=10)
            
    pdf.output(f"PDFs/{filename}.pdf")


import glob
from pathlib import Path
import pandas as pd
from fpdf import FPDF

CURRENCY = "Euros"
filepaths = glob.glob("samples/*.xlsx")

for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    
    # Initialize the PDF file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    # Implement the template to fill with data.
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)
    
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)
    
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    # Set headers 
    columns = [col.replace("_", " ") for col in df.columns]
    
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=columns[0].title(), border=1)
    pdf.cell(w=60, h=8, txt=columns[1].title(), border=1)
    pdf.cell(w=35, h=8, txt=columns[2].title(), border=1)
    pdf.cell(w=30, h=8, txt=columns[3].title(), border=1)
    pdf.cell(w=30, h=8, txt=columns[4].title(), border=1, ln=1)
    
    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add total price row
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(df["total_price"].sum()), border=1, ln=1)
    
    # Add text representation of the total price
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=10, h=20, txt=f"The total due amount is {df['total_price'].sum()} {CURRENCY}", ln=1)

    # Add company and logo
    pdf.set_font(family="Times", size=20, style="B")
    pdf.cell(w=35, h=8, txt=f"PythonHow")
    pdf.image("samples/pythonhow.png", w=10)

    
    # Save the output PDF file.
    pdf.output(f"PDFs/{filename}.pdf")
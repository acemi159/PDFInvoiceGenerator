import glob
from pathlib import Path
import pandas as pd
from fpdf import FPDF


filepaths = glob.glob("samples/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
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
    
    # Save the output PDF file.
    pdf.output(f"PDFs/{filename}.pdf")
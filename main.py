import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob(pathname="invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")

    filename = Path(filepath).stem
    invoice_num, invoice_date = filename.split("-")

    pdf.cell(w=0, h=8, text=f"Invoice nr.{invoice_num}", align="L", ln=1)

    # Convert date string to datetime object
    date_obj = pd.to_datetime(invoice_date, format="%Y.%m.%d")
    # Format date as desired (optional)
    formatted_date = date_obj.strftime("%B %d, %Y")

    pdf.cell(w=0, h=8, text=f"Date: {formatted_date}", align="L")
    pdf.output(f"PDFs/{filename}.pdf")

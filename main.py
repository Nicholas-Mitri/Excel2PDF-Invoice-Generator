import pandas as pd
import glob
from fpdf import FPDF, XPos, YPos
from pathlib import Path

# Get all Excel files from invoices directory
filepaths = glob.glob(pathname="invoices/*.xlsx")


# Process each Excel file and create a PDF invoice
for filepath in filepaths:
    # Read Excel file into DataFrame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Create PDF object
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")

    # Extract invoice number and date from filename
    filename = Path(filepath).stem
    invoice_num, invoice_date = filename.split("-")

    # Add invoice number to PDF
    pdf.cell(
        w=0,
        h=8,
        text=f"Invoice nr.{invoice_num}",
        align="L",
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )

    # Convert date string to datetime object
    date_obj = pd.to_datetime(invoice_date, format="%Y.%m.%d")
    # Format date as desired (optional)
    formatted_date = date_obj.strftime("%B %d, %Y")

    # Add formatted date to PDF
    pdf.cell(
        w=0,
        h=8,
        text=f"Date: {formatted_date}",
        align="L",
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )

    # Add empty line for spacing
    pdf.cell(
        w=0,
        h=8,
        text="",
        align="L",
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )

    # Format column headers
    cols = [col.replace("_", " ").title() for col in df.columns]

    # Set font for table headers
    pdf.set_font(family="Times", size=11, style="B")
    pdf.set_text_color(80, 80, 80)
    # Add table headers
    pdf.cell(w=30, h=8, text=cols[0], border=1)
    pdf.cell(w=50, h=8, text=cols[1], border=1)
    pdf.cell(w=40, h=8, text=cols[2], border=1)
    pdf.cell(w=30, h=8, text=cols[3], border=1)
    pdf.cell(
        w=30,
        h=8,
        text=cols[4],
        border=1,
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )

    # Add table data rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, text=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, text=row["product_name"], border=1)
        pdf.cell(w=40, h=8, text=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, text=str(row["price_per_unit"]), border=1)
        pdf.cell(
            w=30,
            h=8,
            text=str(row["total_price"]),
            border=1,
            new_x=XPos.LMARGIN,
            new_y=YPos.NEXT,
        )

    # Add total row
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, text="", border=1)
    pdf.cell(w=50, h=8, text="", border=1)
    pdf.cell(w=40, h=8, text="", border=1)
    pdf.cell(w=30, h=8, text="", border=1)
    pdf.cell(
        w=30,
        h=8,
        text=str(df["total_price"].sum()),
        border=1,
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )

    # Add total price text
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(
        w=0,
        h=8,
        text=f"The total price is {df["total_price"].sum()}.",
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )

    # Add company info and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, text="PythonHow")
    pdf.image("pythonhow.png", w=10)

    # Save PDF file
    pdf.output(f"PDFs/{filename}.pdf")

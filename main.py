import pandas as pd
import glob
from PIL import Image
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nbr = filename.split("-")[0]
    date = filename.split("-")[1]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice number: {invoice_nbr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header to table
    header = list(df.columns)
    header = [item.replace("_", " ").title() for item in header]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=header[0], border=1)
    pdf.cell(w=70, h=8, txt=header[1], border=1)
    pdf.cell(w=35, h=8, txt=header[2], border=1)
    pdf.cell(w=30, h=8, txt=header[3], border=1)
    pdf.cell(w=30, h=8, txt=header[4], border=1, ln=1)

    # Add rows to table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=35, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    # Total sum row wih sum value of fields
    total_sum = df['total_price'].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=70, h=8, txt=" ", border=1)
    pdf.cell(w=35, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=10, txt=f"The total  price is {total_sum}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=40, txt="Invoice Generator")

    # Logo Image
    pdf.image("thumbnail.png", w=40)

    pdf.output(f"PDFs/{filename}.pdf")



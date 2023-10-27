import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path


filepaths = glob.glob("Invoices/*.xlsx")

for file in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    df = pd.read_excel(file, sheet_name="Sheet 1")
    pdf.add_page()
    filename = Path(file).stem
    invoice_no = filename.split('-')[0]
    date = filename.split('-')[1]

    pdf.set_font(family="arial", style="BU", size=18)
    pdf.cell(w=100, h=10, txt=f"Invoice No. = {invoice_no}", align="L",
             border=0, ln=1)
    pdf.cell(w=100, h=10, txt=f"Date = {date}", align="l",
             border=0, ln=1)
    pdf.ln(10)

    col = df.columns
    col = [item.replace("_", " ").title() for item in col]
    pdf.set_font(family="arial", style="", size=10)
    pdf.set_line_width(0.4)
    pdf.cell(w=31, h=8, txt=col[0], border=1, ln=0, align="L")
    pdf.cell(w=63, h=8, txt=col[1], border=1, ln=0, align="L")
    pdf.cell(w=32, h=8, txt=col[2], border=1, ln=0, align="R")
    pdf.cell(w=32, h=8, txt=col[3], border=1, ln=0, align="R")
    pdf.cell(w=32, h=8, txt=col[4], border=1, ln=1, align="R")

    for index, row in df.iterrows():
        pdf.set_font(family="arial", style="", size=10)
        pdf.set_line_width(0.2)
        pdf.cell(w=31, h=8, txt=str(row["product_id"]), border=1,
                 ln=0, align="L")
        pdf.cell(w=63, h=8, txt=row["product_name"], border=1,
                 ln=0, align="L")
        pdf.cell(w=32, h=8, txt=str(row["amount_purchased"]), border=1,
                 ln=0, align="R")
        pdf.cell(w=32, h=8, txt=str(row["price_per_unit"]), border=1,
                 ln=0, align="R")
        pdf.cell(w=32, h=8, txt=str(row["total_price"]), border=1,
                 ln=1, align="R")

    pdf.cell(w=190, h=8, txt=str(df['total_price'].sum()), border=1,
             ln=1, align="R")
    pdf.ln(20)
    pdf.set_font(family="arial", style="B", size=12)
    pdf.cell(w=190, h=8,
             txt=f"The total due amount is {df['total_price'].sum()} euros",
             border=0, ln=1, align="L")

    pdf.output(f"{filename}.pdf")

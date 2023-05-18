import pandas as pd
import glob as gb
from fpdf import FPDF
from pathlib import Path

filepaths = gb.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    date = filename.split("-")[1]

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr: {invoice_nr}", ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    invoice_columns = list(df.columns)
    invoice_columns = [item.replace("_", " ").title() for item in invoice_columns]
    pdf.set_font(family="Times", style="B", size=12)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=invoice_columns[0], border=1)
    pdf.cell(w=60, h=8, txt=invoice_columns[1], border=1)
    pdf.cell(w=40, h=8, txt=invoice_columns[2], border=1)
    pdf.cell(w=30, h=8, txt=invoice_columns[3], border=1)
    pdf.cell(w=30, h=8, txt=invoice_columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", style="B", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add total_sum
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", style="B", size=12)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=60, h=8, txt=" ", border=1)
    pdf.cell(w=40, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Make a statement
    pdf.set_font(family="Times", style="B", size=12)
    pdf.cell(w=30, h=8, txt=f"The total price is: {total_sum}", ln=1)

    # Add a company name and logo
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=25, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices\*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Display invoice numbers and dates
    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=30, h=10, txt=f"Invoice No. {invoice_nr}", align="L", ln=1, border=0)
    pdf.cell(w=30, h=10, txt=f"Date {date}", align="L", ln=1, border=0)

    # Add a blank row.
    pdf.cell(w=0, h=8, txt="", align="L", ln=1, border=0)

    # Get the data from Excel files.
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Get the headers
    headers = df.columns
    # Trim the header names
    headers = [header.replace("_", " ").title() for header in headers]

    # Add the headers that make the first row of the table
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=headers[0], border=1)
    pdf.cell(w=70, h=8, txt=headers[1], border=1)
    pdf.cell(w=30, h=8, txt=headers[2][:6], border=1)
    pdf.cell(w=30, h=8, txt=headers[3], border=1)
    pdf.cell(w=30, h=8, txt=headers[4], border=1, ln=1)

    # Add the contents data of the invoices
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(0, 0, 0)

        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=row["product_name"], border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Calculate the total prices.
    # total = 0
    # for price in df["total_price"]:
    #     total = total + price

    # Video solution;
    total = df["total_price"].sum()

    # Display the total amount in the row.
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total), border=1, ln=1)
    # Add two blank rows
    pdf.cell(w=0, h=8, txt="", align="L", ln=1, border=0)
    pdf.cell(w=0, h=8, txt="", align="L", ln=1, border=0)

    # Display the total due amount statement.
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=10, txt=f"The total due amount is {total} yen.",
             align="L", ln=1, border=0)

    # Display the company name and logo.
    pdf.cell(w=27, h=10, txt="LoveCats", align="L", ln=0, border=0)
    pdf.image("cat_eyes.png", w=0, h=10)

    # Output them in PDFs.
    pdf.output(f"PDFs\{filename}.pdf")










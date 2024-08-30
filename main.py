import glob
import pandas as pd
from fpdf import FPDF


def get_width(value):
    return pdf.get_string_width(str(value))


def write_cell(text, w=20, h=12, align='L', border=1, ln=0, index=-1, ):
    match index:
        case 1:
            w += 30
        case 2:
            w += 10
        case 4:
            ln = 1

    pdf.cell(w=w, h=h, txt=str(text), align=align, border=border,
             ln=ln)


filepaths = glob.glob('invoices/*.xlsx')
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    df_list = df.values.tolist()
    total_price = sum([price for price in df['total_price']])

    pdf = FPDF(orientation="P", unit="mm", format='A4')
    pdf.add_page()

    # create title
    pdf.set_font(family="Times", style="B", size=24)
    pdf.set_text_color(100, 100, 100)
    write_cell(f"Invoice nr.: {filepath[9:14]}", w=50, ln=1, border=0)
    write_cell(f"Date : "
               f"{filepath[15:].strip('.xlsx').replace('.', '-')}",
               w=50, ln=1, border=0)
    pdf.ln(20)

    # Create table header
    pdf.set_font(family="Times", size=8)
    columns = df.columns.to_list()
    for index, column in enumerate(columns):
        write_cell(column.replace("_", " ").title(), index=index)

    # Create table content
    for row in df_list:
        for index, item in enumerate(row):
            write_cell(item, index=index)

    # Create last row of table with total sum
    for index, item in enumerate(["", "", "", "", total_price]):
        write_cell(item, index=index)
    pdf.ln(30)

    # Create footer with total price
    pdf.set_font(family="Times", style="B", size=16)
    write_cell(f"The total due amount is: {total_price} euros", border=0)
    pdf.output(f"output/invoice_{filepath[9:14]}.pdf")

import pandas as pd
import fpdf
import glob
import pathlib
import openpyxl
import os
filepaths = glob.glob("invoices_excel/*.xlsx")

print(filepaths)
for filepath in filepaths:
    df=pd.read_excel(filepath)
    # sheet_name because an excel file can have multiple sheets

    # extracting invoice number
    filename = pathlib.Path(filepath).stem
    invoice_no = filename.split("-")[0]

    # create a pdf corresponding to each excel invoice
    pdf = fpdf.FPDF(orientation="P", unit='mm', format="A4")
    pdf.add_page()
    pdf.set_font(family="Times",style="B", size=16)
    pdf.cell(w=0, h=20, txt=f"Invoice No# {invoice_no}")

    # create directory after checking its existence
    if os.path.isdir("PDFs") == False:
        os.mkdir("PDFs")

    pdf.output(f"PDFs/{filename}.pdf")



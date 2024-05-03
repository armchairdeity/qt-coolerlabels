"""
Module Name: labelizer
Project: Label Generation
Copyright (c) 2023 Amazon

This module contains the main logic for generating labels
from product data. It reads Excel files, generates individual
barcode images, defines the label template, populates it with
product data and barcodes, and outputs the finished labels PDF.

The functions getProductDatafile(), generateBarcodes(),
getSize(), and drawLabel() are imported from the funcs.py
module to handle preprocessing tasks.

This source code is provided under an MIT license. See
LICENSE file for more details.

Here is an explanation of the selected code block:

The code is importing necessary libraries and functions for generating labels from product data.

It starts by cleaning up leftover files from previous runs, to ensure a clean workspace.

It then gets a list of Excel files to process from the prod_docs folder.

The main logic loop iterates through each Excel file:

It calls getProductDatafile() to read the file and return a list of product data
It calls generateBarcodes() to create individual barcode images
It sets up the label template specification using the ReportLabel library
It draws each product's barcode onto a label, getting the image path and print quantity
Any errors are caught and output
Once complete, it saves the finished PDF label sheet
The key steps are:
	preparing the data,
	generating barcode images,
	defining the label template,
	populating it with product-specific barcodes,
	and outputting the finished labels PDF.
"""

import os
import sys
import trace
import labels
import glob
import traceback

from funcs import getProductDatafile
from funcs import generateBarcodes
from funcs import getSize
from funcs import drawLabel

# clean up leftover files from previous runs
if os.path.exists("./barcodes"):
    for png in glob.glob("./barcodes/*.png"):
        os.remove(png)
if os.path.exists("./finished"):
    for pdf in glob.glob("./finished/*.pdf"):
        os.remove(pdf)

# get the list of Excel files
files = [
    f
    for f in os.listdir("./prod_docs")
    if os.path.splitext(f)[1] == ".xlsx"
    if f[0] != "~"
]

# loop the list of Excel files
for file in files:
    print("")
    print("Filename: ", file)
    try:
        # read Excel file and return a list of 3-element lists
        productData = getProductDatafile(f"./prod_docs/{file}")
        prodName = os.path.split(os.path.splitext(file)[0])[1]

        # Generate the barcode images 1 at a time
        generateBarcodes(productData)

        spec = labels.Specification(
            sheet_width=getSize(8.5),
            sheet_height=getSize(11),
            top_margin=getSize(0.5),
            bottom_margin=getSize(0.5),
            left_margin=getSize(0.1875),
            columns=3,
            rows=10,
            row_gap=getSize(0),
            column_gap=getSize(0.125),
            label_width=getSize(2.625),
            label_height=getSize(1),
        )
        sheet = labels.Sheet(spec, drawLabel, border=False)

        print("\nAdding UPC barcodes to labels PDF:", end="\n")

        for p in productData:
            try:  # "PrintCount", "UPC", "Name"
                p[2] = p[2].replace("\n", " ")
                print(*p, end="\n", sep=" | ")
                imgPath = os.path.abspath(f"barcodes/{p[1]}.png")
                # Adds the new label to the sheet and includes the # of labels to print
                # from the first column of the spreadsheet
                sheet.add_label(imgPath, int(p[0]))
            except Exception as e:
                break
                print("\nA product data exception occurred!\n\nException text reads:")
                print(e, end="\n")

        print(f"Saving PDF: {prodName}")
        sheet.save(f"./finished/{prodName}.pdf")
    except Exception as e:
        print("\nA file exception exception occurred!\nException text reads:")
        print(traceback.format_exc())

"""
Module Name: funcs
Project: Barcode Label Generation
Copyright (c) 2023 Amazon

This module contains helper functions used in the label generation
process. It handles reading product data from Excel files, generating
individual barcode images, setting label dimensions, and drawing 
barcodes onto labels.

These functions are imported and used by the labelizer.py module,
which contains the main logic for generating labels from product data.

Functions:
  - getProductDatafile(): Reads product data from an Excel file
  - generateBarcodes(): Creates barcode images from product data  
  - getSize(): Converts dimensions between inches and mm
  - drawLabel(): Draws a barcode image onto a label

This source code is provided under an MIT license. See 
LICENSE file for more details.
"""
import barcode
import pandas as pd
import os
from barcode.writer import ImageWriter
from reportlab.graphics import shapes

# Filename MUST be an Excel file, reads it and returns the contents in list of lists
def getProductDatafile(fileName: str) -> list:
	with open(fileName) as products:
		retVal = [r.split("\t") for r in pd.read_excel(fileName, header=None, usecols=[0, 1, 2], dtype="string", names=["PrintCount", "UPC", "Name"]).to_csv(sep="\t", index=False, header=False).split("\n")]
		print(*retVal)
		return retVal
	
# callback for the Sheet(), applies an Image to the label
def drawLabel(label:shapes.Drawing, width:int, height:int, obj:str):
	img = shapes.Image(x=20, y=0, width=width-60, height=height, path=str(obj))
	label.add(img)

# cleans up generated CSV files between runs
def clearFiles_csv(folderPath):
	[os.remove(csv) for csv in os.listdir(folderPath) if os.path.splitext(csv) == ".csv"]

# Generate the barcode images 1 at a time, looping over the list that 
# gets returned from a call to getProductData()
def generateBarcodes(prodData: list):

	print("Generating bar codes!\n\n")
	for a in prodData:
		
		print(*a,end="\n", sep=" : ")
		
		# break out of the loop if the list item has less than the 3 expected elements
		# this can happen sometimes on the last iteration of the loop if your lines
		# are pulled from a data file that ends in a newline, yanno, like MOST of them. ;)
		if len(a) < 3:
			break

		imageWriterOpts = dict(
			format="png",
			module_height=10,
			# Changes the font size depending on the length of the product name (a[2])
			font_size=10 if len(a[2]) <= 20 else 8 if len(a[2]) >= 24 else 9,
			center_text=True,
			text_distance=4 if len(a[2]) <= 20 else 3 if len(a[2]) >= 26 else 5
		)
		try:
			# if the product name is longer than 26 characters, break it up into multiple lines
			if len(a[2]) >= 26:
				# replace "/" with "/ " so it can line-break after a slash
				a[2]=a[2].replace("/"," ")
				# create a list out of the letters in the product name
				chars=[c for c in a[2]]
				# this is complex... OK, here we go:
				# change the value of the first space after the first 33% of the list from a space to a "\n"
				chars[a[2].find(" ",int(len(chars)*.33))]="\n"
				# now reassemble the string by doing a join() call against an empty string, thus retaining the original
				# letters and spaces, as well as the new newline character
				a[2]="".join(chars)
			
			# now generate the bar code
			barcode.generate('upc', a[1], ImageWriter(), f"./barcodes/{a[1]}", text=f"{a[2]}", writer_options=imageWriterOpts)

			# now reset the newline in the product name to a space so that it doesn't mess up thee output of the 
			# pdf generation process.
			a[2].replace("\n"," ")

		except IndexError as i:
			break
		except Exception as e:
			print(*a)
			print(*e)

	print(f"Generation of barcode images is complete!")

# converts decimal inches to milimeters for setting up the Specification
def getSize(size: float) -> float:
	return size*25.4

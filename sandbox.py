import os
import sys
import labels

from funcs import getProductDatafile
from funcs import generateBarcodes
from funcs import getSize
from funcs import drawLabel

# if (os.path.exists('./barcodes')):
# 	for i in glob.glob('./barcodes/*.png'):
# 		os.remove(i)


file = [
    f
    for f in os.listdir("./prod_docs")
    if os.path.splitext(f)[1] == ".xlsx"
    if f[0] != "~"
][1]

print("")
print("Filename: ", file)
try:
    productData = getProductDatafile(f"./prod_docs/{file}")
    productData.sort()
    prodName = os.path.split(os.path.splitext(file)[0])[1]

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
    # except Exception as e:
    # 	print(e)
    prevP = []

    for p in productData:
        try:  # "PrintCount", "UPC", "Name"
            print(*p, end="\n", sep=" | ")
            imgPath = os.path.abspath(f"barcodes/{p[1]}.png")
            sheet.add_label(imgPath, int(p[0]))
            prevP.append(p)
        except Exception as e:
            # break
            print("\nA product data exception occurred!\n\nException text reads:")
            print(prodName, end="\n")
            print(prevP, end="\n")
            print(e, end="\n")

    print(f"Saving PDF: {prodName}")
    sheet.save(f"./finished/{prodName}.pdf")
except Exception as e:
    print("\nA file exception exception occurred!\nException text reads:")
    print(e)

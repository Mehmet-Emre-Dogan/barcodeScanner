# REFERENCES
# https://www.geeksforgeeks.org/how-to-make-a-barcode-reader-in-python/
# Importing library
import cv2
from pyzbar.pyzbar import decode
import os
import re
from msvcrt import getch
import pandas as pd
import datetime

VALID_EXTENSIONS = (".JPG", ".JPEG", ".jpg", ".jpeg")
  
# Make one method to decode the barcode
def barcodeReader(image):
     
    # read the image in numpy array using cv2
    img = cv2.imread(image)
      
    # Decode the barcode image
    detectedBarcodes = decode(img)
      
    # If not detected then print the message
    if not detectedBarcodes:
        print("Barcode Not Detected or your barcode is blank/corrupted!")
    else:
       
          # Traverse through all the detected barcodes in image
        for barcode in detectedBarcodes: 
           
            # Locate the barcode position in image
            (x, y, w, h) = barcode.rect
             
            # Put the rectangle in image using
            # cv2 to heighlight the barcode
            cv2.rectangle(img, (x-10, y-10),
                          (x + w+10, y + h+10),
                          (255, 0, 0), 2)
             
            if barcode.data!="":
               
            # Print the barcode data
                return barcode.data.decode("utf-8") 
                # print(barcode.type)
                 
# https://stackoverflow.com/questions/4836710/is-there-a-built-in-function-for-string-natural-sort
def naturalSort(arr): 
    convert = lambda text: int(text) if text.isdigit() else text.lower()
    alphanum_key = lambda key: [convert(c) for c in re.split('([0-9]+)', key)]
    return sorted(arr, key=alphanum_key)

def runreader(path):
    myDict = {"filenames": [],  "barcodes": [], "itemCounts": []}
    customDTStart = datetime.datetime.now().strftime('%Y-%m-%d_%H.%M.%S')
    # Scan for the images
    files = naturalSort(os.listdir(path + "input"))
    # print(files)
    cou = 0
    for i, item in enumerate(files):
      if str(item).lower().endswith(VALID_EXTENSIONS):
        decodeResult = barcodeReader(path + "input\\" + item)
        print(f"{str(cou+1).zfill(2).rjust(3)}- Image found: {item} #Barcode: {decodeResult}")
        myDict["filenames"].append(item)
        myDict["barcodes"].append(decodeResult)
        myDict["itemCounts"].append(None)
        cou += 1

    #################################################################################################
    """Excel Writing"""
    df = pd.DataFrame(myDict)
    df["Number"] = df.index + 1
    df = df[["Number", "filenames", "barcodes", "itemCounts"]]
    print(df.to_string(index=False))

    # Please see the below sources for further information
    # https://stackoverflow.com/questions/22831520/how-to-do-excels-format-as-table-in-python
    # https://xlsxwriter.readthedocs.io/example_pandas_table.html
    # https://xlsxwriter.readthedocs.io/working_with_tables.html
    # https://stackoverflow.com/questions/17326973/is-there-a-way-to-auto-adjust-excel-column-widths-with-pandas-excelwriter

    print("Writing to excel file...")
    writer = pd.ExcelWriter(f"Barcodes_{customDTStart}.xlsx", engine='xlsxwriter')
    df.to_excel(writer, sheet_name="Barcodes", index=False, startrow=1, header=False)
    workbook = writer.book
    worksheet = writer.sheets['Barcodes']
    (rowCou, ColCou) = df.shape
    columnSettings = [{'header': column} for column in df.columns]
    worksheet.add_table(0, 0, rowCou, ColCou - 1, {'columns': columnSettings, 'style': 'Table Style Medium 4'})
    for i, column in enumerate(df.columns):
        colLen = df[column].astype(str).str.len().max()
        colLen = max(colLen, len(column)) # colLen is the maximum length of the rows in this column. And len(column) is the length of the header of this column
        worksheet.set_column(first_col=i, last_col=i, width=colLen)

    writer.save()
    #################################################################################################
 
if __name__ == "__main__":
    currDir = ".\\" # Get the current directory
    runreader(currDir)
    print("Writing completed. Press any key to exit...")
    getch()
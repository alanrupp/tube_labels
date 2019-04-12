#!/home/alanrupp/anaconda3/bin/python

## Create script to generate tube label sheet (12x16 labels)
import numpy as np
import pandas as pd
import xlsxwriter
import re
import argparse

# parse command line arguments
parser = argparse.ArgumentParser(description="Generate a sheet of tube labels")
parser.add_argument("--start", type=str)
parser.add_argument("--size", type=str, default='1/2"')
parser.add_argument("--outfile", type=str, default="tube_labels.xlsx")
args = parser.parse_args()

# grab start value and parse letter and number
spaceOne = args.start
letter = re.findall('[A-Za-z]+', spaceOne)[0]
if re.search('[0-9]+', spaceOne):
    number = int(re.findall('[0-9]+', spaceOne)[0])
else:
    number = list()

# generate matrix of values
labels = list()
if isinstance(number, int):
    while number <= 1000 and len(labels) < 192:
        value = letter + str(number)
        number += 1
        labels.append(value)
else:
    labels = [letter] * 192

# put into pandas dataframe
labels = np.array(labels).reshape(16,12)
labels = pd.DataFrame(labels)

# prepare Excel writer
writer = pd.ExcelWriter(path=args.outfile, engine='xlsxwriter')
workbook = writer.book

# adjust font size for each label
labelSize = args.size
if labelSize == '1/2"':
    fontSize = 11
elif labelSize == '3/8"':
    fontSize = 10

cell_format = workbook.add_format({'text_wrap': True, 'align': 'center',
    'valign': 'vcenter', 'font_name': 'Arial', 'font_size': fontSize})

# write to Excel
labels.to_excel(writer, sheet_name='Sheet1', header=False, index=False)
worksheet = writer.sheets['Sheet1']
worksheet.set_column('A:L', 5.7, cell_format)
for i in range(16):
    worksheet.set_row(i, 45)
worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)

workbook.close()

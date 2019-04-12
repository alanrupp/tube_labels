#!/home/alanrupp/anaconda3/bin/python

## Create script to generate tube label sheet(s) (12x16 labels)
import numpy as np
import pandas as pd
import xlsxwriter
import re
import argparse

# parse command line arguments
parser = argparse.ArgumentParser(description="Generate a sheet of tube labels")
parser.add_argument('--start', help='specify a start value', type=str)
parser.add_argument('--end', help='optional end value (default makes 1 sheet)',\
                    type=str, default=False)
parser.add_argument("--size", help='sticker size in inches (default is 1/2")', \
                    type=str, default='1/2"')
parser.add_argument("--outfile", help='name of output file (default is tube_labels.xlsx)', \
                    type=str, default='tube_labels.xlsx')
args = parser.parse_args()

# funciton to grab letter and number
def parse_label(label):
    letter = re.findall('^[A-Za-z]+', label)
    if len(letter) < 1:
        letter = ""
    else:
        letter = letter[0]
    # grab number
    if re.search('[0-9]+$', label):
        number = int(re.findall('[0-9]+$', label)[0])
    else:
        number = ""
    return(letter, number)

# find first letter and/or number
first_letter, first_number = parse_label(args.start)

# find last letter and/or number
if not args.end:
    last_number = first_number + 191
    last_letter = first_letter + str(last_number)
else:
    last_letter, last_number = parse_label(args.end)
    if last_letter != first_letter:
        print("\nError: Start and End values are not compatible")
        exit()
    else:
        letter = first_letter
    if type(last_number) != type(first_number):
        print("\nError: Start and End values are not compatible")
        exit()

# generate all the values
labels = list()
number = first_number
if isinstance(number, int):
    while number <= 1000 and number <= last_number:
        value = letter + str(number)
        number += 1
        labels.append(value)
else:
    labels = [letter] * 192

# - put into pandas dataframe  ------------------------------------------------
total_labels = len(labels)
if total_labels % (12*16) != 0:
    while total_labels % (12*16) != 0:
        labels.append('')
        total_labels += 1
rows = int(total_labels / 12)

labels = np.array(labels).reshape(rows, 12)
labels = pd.DataFrame(labels)

# - prepare Excel writer ------------------------------------------------------
total_sheets = int(len(labels) / 16)
writer = pd.ExcelWriter(path=args.outfile, engine='xlsxwriter')
workbook = writer.book

# adjust font size for each label
if args.size == '1/2"':
    fontSize = 11
elif args.size == '3/8"':
    fontSize = 10

# add formatting for appropriate font sizes and centering
cell_format = workbook.add_format({'text_wrap': True, 'align': 'center',\
    'valign': 'vcenter', 'font_name': 'Arial', 'font_size': fontSize})

# write to Excel
for sheet in range(total_sheets):
    labels_write = labels.iloc[sheet*16: (sheet+1)*16]
    labels_write.to_excel(writer, sheet_name='Sheet' + str(sheet), \
                     header=False, index=False)
    worksheet = writer.sheets['Sheet' + str(sheet)]
    worksheet.set_column('A:L', 5.7, cell_format)
    for i in range(16):
        worksheet.set_row(i, 45)
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)

workbook.close()

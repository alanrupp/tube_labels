#!/home/alanrupp/anaconda3/bin/python

## Create script to generate tube label sheet(s) (12x16 labels)
import numpy as np
import pandas as pd
import xlsxwriter
import re


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


# generate all the values
def make_labels(first_letter, first_number):
    # find last letter and/or number
    if not args.end and not args.total:
        last_number = first_number + 191
        last_letter = first_letter + str(last_number)
        letter = first_letter
    elif args.total:
        last_number = first_number + args.total
        last_letter = first_letter + str(last_number)
        letter = first_letter
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
    # put into a list
    labels = list()
    number = first_number
    if isinstance(number, int):
        while number <= 1000 and number <= last_number:
            value = letter + str(number)
            number += 1
            labels.append(value)
    else:
        labels = [letter] * 192
    return labels

# put into pandas dataframe
def labels_to_df(labels):
    total_labels = len(labels)
    if total_labels % (12*16) != 0:
        while total_labels % (12*16) != 0:
            labels.append('')
            total_labels += 1
    rows = int(total_labels / 12)
    labels = np.array(labels).reshape(rows, 12)
    labels = pd.DataFrame(labels)
    return labels

# prepare Excel writer
def write_excel(labels, size, output):
    total_sheets = int(len(labels) / 16)
    writer = pd.ExcelWriter(path=output, engine='xlsxwriter')
    workbook = writer.book

    # adjust font size for each label
    if size == "big":
        fontSize = 11
    elif size == "small":
        fontSize = 8

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
            if size == "big":
                worksheet.set_margins(left=0.52, right=0.52, top=0.52, bottom=0.52)
            elif size == "small":
                worksheet.set_margins(left=0.62, right=0.69, top=0.62, bottom=0.62)

    workbook.close()


# - Run -----------------------------------------------------------------------
if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="Generate a sheet of tube labels")
    parser.add_argument('start', help='specify a start value', type=str)
    parser.add_argument('--end', help='optional end value (default makes 1 sheet)',\
                        type=str, default=False)
    parser.add_argument("--size", help='sticker size: "big" or "small" (default is big)', \
                        type=str, default='big')
    parser.add_argument("--outfile", help='name of output file (default is tube_labels.xlsx)', \
                        type=str, default='tube_labels.xlsx')
    parser.add_argument('--total', help='total stickers', default=192, type=int)
    parser.add_argument('--text', help='treat all inputs as text', action='store_true')
    args = parser.parse_args()

    # find first letter and/or number
    first_letter, first_number = parse_label(args.start)
    if (args.text):
        first_letter = args.start
        first_number = ""
        labels = [args.start] * args.total
    else:
        first_letter, first_number = parse_label(args.start)
        labels = make_labels(first_letter, first_number)

    # make pandas dataframe
    labels = labels_to_df(labels)

    # write to Excel
    write_excel(labels, args.size, args.outfile)

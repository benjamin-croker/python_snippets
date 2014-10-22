# Takes a large CSV file and splits it into individual worksheets of an Excel workbook

import csv
import itertools
import sys

from openpyxl import Workbook


def split_csv_generator(reader, split_size):
    header = reader.next()
    while True:
        block = [b for b in itertools.islice(reader, split_size)]
        if len(block) > 0:
            yield header, block
        else:
            break
    
    
def split_into_workbook_sheets(split_gen, ws_prefix):
    wb = Workbook()
    for i, (header, row) in enumerate(split_gen):
        # there will be one sheet when the workbook is created
        if i == 0:
            ws = wb.get_active_sheet()
        else:
            ws = wb.create_sheet()
        ws.title = '{}_{}'.format(ws_prefix, i)
        ws.append(header)
        for r in row:
            ws.append(r)
            
    return wb
    
        
def split(in_filename, out_filename, split_size, ws_prefix):
    reader = csv.reader(open(in_filename, 'rb'))
    split_gen = split_csv_generator(reader, split_size)
    
    wb = split_into_workbook_sheets(split_gen, ws_prefix)
    wb.save(out_filename)
    
    
if __name__ == '__main__':

    if len(sys.argv) != 5:
        print("Usage:\tsplit_csv.py input_file output_file block_size worksheet_prefix")
    
    else:
        split(sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4])
        

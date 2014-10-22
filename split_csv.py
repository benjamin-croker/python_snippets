# Takes a large CSV file and splits it into individual Excel workbooks
 
import csv
import itertools
import sys
import os
 
from openpyxl import Workbook
 
 
def split_csv_generator(reader, split_size):
    header = reader.next()
    while True:
        block = [b for b in itertools.islice(reader, split_size)]
        if len(block) > 0:
            yield header, block
        else:
            break
    
    
def split_into_workbook_sheets(split_gen, out_dir, ws_prefix):
    for i, (header, row) in enumerate(split_gen):
        # there will be one sheet when the workbook is created
        wb = Workbook()
        ws = wb.get_active_sheet()
        ws.title = '{}_{}'.format(ws_prefix, i)
        ws.append(header)
        for r in row:
            ws.append(r)
        wb.save(os.path.join(out_dir, '{}.xlsx'.format(ws.title)))
    
        
def split(in_filename, out_dir, split_size, ws_prefix):
    reader = csv.reader(open(in_filename, 'rb'))
    split_gen = split_csv_generator(reader, split_size)
    split_into_workbook_sheets(split_gen, out_dir, ws_prefix)
    
    
if __name__ == '__main__':
 
    if len(sys.argv) != 5:
        print("Usage:\tsplit_csv.py input_file output_directory block_size worksheet_prefix")
    
    else:
        split(sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4])

import xlsxwriter
import sys
import pandas as pd
import os


if __name__ == '__main__':
   filename = sys.argv[1]
   base_filename, file_extension = os.path.splitext(filename)
   output_filename = base_filename+"-formatted" + file_extension
   print(f"Reading file: {filename}")
   df = pd.read_excel(filename, header=None)
   wb = xlsxwriter.Workbook(output_filename)
   ws = wb.add_worksheet()
   for i, row in enumerate(df.values):
      ws.write_row(i, 0, row)
   f1 = wb.add_format({'bg_color': '#D9D9D9', 'font_color': 'red'})
   ws.conditional_format(
      'A1:B10',{
         'type':'cell', 'criteria':'<', 'value':50, 'format':f1
      } 
   )
   f2 = wb.add_format({'bg_color': '#D9D9D9', 'font_color': 'blue'})
   ws.conditional_format(
      'A1:B10',{
         'type':'cell', 'criteria':'>', 'value':50, 'format':f2
      } 
   )
   f3 = wb.add_format({'bg_color': '#D9D9D9', 'font_color': 'green'})
   ws.conditional_format(
      'A1:B10',{
         'type':'cell', 'criteria':'=', 'value':50, 'format':f3
      } 
   )
   wb.close()
   print(f"Done. File Saved to: {output_filename}")

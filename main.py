from docxtpl import DocxTemplate
import pandas
import argparse
import os

parser = argparse.ArgumentParser(description='Filling document template')
parser.add_argument('template', type=str, help='document template file')
parser.add_argument('datafile', type=str, help='file with data to fill template')
parser.add_argument('-column', type=str, help='column name to append to final filenames (default: row index)')
args = parser.parse_args()

excel_data_df = pandas.read_excel(args.datafile)
for row in excel_data_df.index:
    doc = DocxTemplate(args.template)
    filename, file_extension = os.path.splitext(args.template)
    context = {}
    for col in excel_data_df.columns:
        context[col] = excel_data_df.loc[excel_data_df.index[row], col]
    doc.render(context)

    index = str(row)
    if args.column is not None:
        index = excel_data_df.loc[excel_data_df.index[row], args.column]

    doc.save(filename + " " + index + file_extension)

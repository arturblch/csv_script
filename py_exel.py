import csv
import xlwt
import os
import argparse
import sys


def check_width(sheet, i, data):
    cwidth = sheet.col(i).width
    if (len(data) * 270) > cwidth:
        sheet.col(i).width = (len(data) * 270)


parser = argparse.ArgumentParser(description='Transform csv to exel file')
parser.add_argument(
    'filename',
    nargs='?',
    type=str,
    help='File name for script',
    default='test.csv')

args = parser.parse_args()

if not os.path.isfile(args.filename):
    print('File not found')
    sys.exit(0)

wb = xlwt.Workbook()
ws = wb.add_sheet('Лист 1')
head = [
    'Номер измерения.', 'Код Атт.', ' Код ФВ.',
    'Затухание по ампл. (ном.), дБ.', 'Сдвииг по фазе (ном.), град.',
    'Затухание по ампл. (изм.), дБ.', 'Сдвииг по фазе (изм.), град.'
]
for col, word in enumerate(head):
    check_width(ws, col, word)
    ws.write(0, col, word)

csv_data = []
with open('test.csv', newline='') as csvfile:
    csv_reader = csv.reader(csvfile, delimiter=',', quotechar='|')
    csv_data = list(csv_reader)

for row in range(len(csv_data)):
    num_att = row % 64
    num_phs = row // 64 % 64
    ws.write(row + 1, 0, row + 1)
    ws.write(row + 1, 2, str(bin(num_phs)[2:].rjust(8, '0')))
    ws.write(row + 1, 1, str(bin(num_att)[2:].rjust(8, '0')))
    ws.write(row + 1, 3, num_att * -0.5)
    ws.write(row + 1, 4, num_phs * 5.625)
    ws.write(row + 1, 5, csv_data[row][0])
    ws.write(row + 1, 6, csv_data[row][1])

os.makedirs('formated', exist_ok=True)
wb.save('formated/{}.xls'.format(args.filename[:-4]))
print('{}.xls Build'.format(args.filename[:-4]))

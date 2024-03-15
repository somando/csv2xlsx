import sys, csv, chardet
from openpyxl import Workbook, load_workbook

args = sys.argv

if len(args) != 2:
    print('使い方: python convert.py <csvファイル名>')
    sys.exit(1)

try:
    with open(args[1], 'rb') as file:
        result = chardet.detect(file.read())
        with open(args[1], 'r', encoding=result['encoding']) as file:
            csv_data = []
            reader = csv.reader(file)
            for row in reader:
                csv_data.append(row)
    
    print(f'{args[1]} を読み込みました')

except FileNotFoundError:
    print(f'{args[1]} が見つかりません')
    sys.exit(1)


wb = Workbook()
wb.save('output.xlsx')

wb = load_workbook('output.xlsx')

wb.create_sheet(title='Sheet', index=0)
wb.remove(wb['Sheet'])

ws = wb['Sheet']

for i, row in enumerate(csv_data):
    for j, column in enumerate(row):
        ws.cell(row=i+1, column=j+1).value = column

wb.save('output.xlsx')
print(f'output.xlsx を保存しました')

import xlrd

wb = xlrd.open_workbook(r'D:\excel merger\DESHMUKH SURYAKANT NARAYANRAO.xls', formatting_info=True)
sheet = wb.sheet_by_index(0)

print(f'Template: {sheet.nrows} rows x {sheet.ncols} cols')
print('=' * 80)

for r in range(min(30, sheet.nrows)):
    vals = []
    for c in range(min(20, sheet.ncols)):
        v = sheet.cell(r, c).value
        if v:
            vals.append(f'[Col{c}] {str(v)[:30]}')
    if vals:
        print(f'Row {r}: {vals}')

print('=' * 80)
print('\nMerged cells:')
for merge in sheet.merged_cells:
    print(f'  Rows {merge[0]}-{merge[1]}, Cols {merge[2]}-{merge[3]}')

import openpyxl
from openpyxl.styles import Font, Alignment
import os

end = 255
start_A = 0
end_A = 1

if not os.path.exists('RGBA_Excels'):
    os.makedirs('RGBA_Excels')

for a in range(start_A, end_A + 1):

    folder_name = f'RGBA_Excels/{a}'
    os.makedirs(folder_name, exist_ok=True)

    for r in range(0, end + 1):

        file_name = f'{folder_name}/RGBA_{r}.xlsx'
        wb = openpyxl.Workbook()

        default_sheet = wb.active
        wb.remove(default_sheet)

        for g in range(0, end + 1):

            sheet = wb.create_sheet(title=f'{g}')

            sheet['A1'].value = 'R'
            sheet['B1'].value = 'G'
            sheet['C1'].value = 'B'
            sheet['D1'].value = 'A'

            bold_font = Font(bold=True)
            center_alignment = Alignment(horizontal='center')
            for col in ['A', 'B', 'C', 'D']:
                cell = sheet[f'{col}1']
                cell.font = bold_font
                cell.alignment = center_alignment

            for b in range(0, end + 1):

                row = (r, g, b, a)
                sheet.append(row)

                print(f'{r},{g},{b},{a}')

        wb.save(file_name)
        print(f'Excel file "{file_name}" created successfully for alpha (a): {a}, red (r): {r}.')

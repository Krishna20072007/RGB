import openpyxl
from openpyxl.styles import Font, Alignment
import os

# Define the alpha value
a = 128

end = 255

# Create the Excels folder if it doesn't exist
if not os.path.exists(f'Excels/{a}'):
    os.makedirs(f'Excels/{a}')

for r in range(0, end+1):
    # Create a new Excel file
    wb = openpyxl.Workbook()
    
    # Remove the original sheet
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    for g in range(0, end+1):
        # Create a new sheet for each g value
        sheet = wb.create_sheet(title=f'{g}')
        
        # Set the column headers
        sheet['A1'].value = 'R'
        sheet['B1'].value = 'G'
        sheet['C1'].value = 'B'
        sheet['D1'].value = 'A'
        
        # Format the column headers
        bold_font = Font(bold=True)
        center_alignment = Alignment(horizontal='center')
        for col in ['A', 'B', 'C', 'D']:
            cell = sheet[f'{col}1']
            cell.font = bold_font
            cell.alignment = center_alignment
        
        for b in range(0, end+1):
            # Populate the data in each row
            row = (r, g, b, a)
            sheet.append(row)
        
        print(f'Appended g = {g} to sheet {g}, book {r}')
    
    # Save the Excel file
    file_name = f'Excels/{a}/RGB_{r}.xlsx'
    wb.save(file_name)
    print(f'Excel file "{file_name}" created successfully.')

from openpyxl import load_workbook
import os
import shutil

data_file = 'Спис.xlsx'
template_file = 'Шабл.xlsx'
output_folder = 'Акты_скрытых_работ'

os.makedirs(output_folder, exist_ok=True)

wb_data = load_workbook(data_file)
ws_data = wb_data['Лист1']

wb_template = load_workbook(template_file)
ws_template = wb_template['Лист1']

for row in ws_data.iter_rows(min_row=2, values_only=True):
    number, description, beton = row

    new_file = os.path.join(output_folder, f"Акт_{number}.xlsx")
    shutil.copy(template_file, new_file)

    wb_new = load_workbook(new_file)
    ws_new = wb_new['Лист1']

    for cell in ws_new['1:100']:
        for c in cell:
            if c.value and isinstance(c.value, str):
                c.value = c.value.replace('[номер акта]', str(number))

    beton_lines = str(beton).split('\n') if beton else ['']
    target_row_beton = 4

    for i, line in enumerate(beton_lines):
        if i > 0:
            ws_new.insert_rows(target_row_beton + i)
        ws_new[f"A{target_row_beton + i}"].value = line
        for cell in ws_new[target_row_beton + i]:
            if cell.value and isinstance(cell.value, str):
                cell.value = cell.value.replace('[бетон]', line)

    wb_new.save(new_file)

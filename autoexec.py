from openpyxl import load_workbook
from datetime import date
import os
import shutil

data_file = 'я. Бетон (Июнь).xlsx'  # Путь к файлу с данными
template_file = 'Шаблон.xlsx'  # Путь к шаблону для актов
output_folder = 'Акты_скрытых_работ'  # Название для папки под акты

os.makedirs(output_folder, exist_ok=True)  # Создаем папку для выходных файлов

wb_data = load_workbook(data_file)
ws_data = wb_data['Бетон для АОСР']

wb_template = load_workbook(template_file)
ws_template = wb_template['АОСР бетон']

for row in ws_data.iter_rows(min_row=3, values_only=True):
    '''Сохраняем данные из файла в exel в переменные. '''

    id = row[0]
    act_number = str(row[1])
    work_name = str(row[2])
    start_date = row[3].strftime('%d.%m.%Y') if isinstance(row[3], date) else str(row[3] or '')
    end_date = row[4].strftime('%d.%m.%Y') if isinstance(row[4], date) else str(row[4] or '')
    concrete_type = str(row[5])
    mixture_number = str(row[7])
    mixture_date = row[9].strftime('%d.%m.%Y') if isinstance(row[9], date) else str(row[9] or '')
    lab_uzk = str(row[10]) if row[10] else ''
    lab_k = str(row[11]) if row[11] else ''
    lab_date = row[12].strftime('%d.%m.%Y') if isinstance(row[12], date) else str(row[12] or '')
    code = str(row[13])
    agreement_date = row[14].strftime('%d.%m.%Y') if isinstance(row[14], date) else str(row[14] or '')

    new_file = os.path.join(output_folder, f'Акт_№{id}.xlsx')
    shutil.copy(template_file, new_file)

    wb_new = load_workbook(new_file)
    ws_new = wb_new['АОСР бетон']

    replacements = {
        '[№ акта]': act_number,
        '[Наименование работ]': work_name,
        '[Дата начала работы]': start_date,
        '[Дата окончания работы]': end_date,
        '[Тип бетонной смеси]': concrete_type,
        '[Смесь №]': mixture_number,
        '[Смесь Дата]': mixture_date,
        '[Лаба УЗК]': lab_uzk,
        '[Лаба К]': lab_k,
        '[Лаба Дата]': lab_date,
        '[Шифр]': code,
        '[Согл Дат]': agreement_date
    }

    for cell in ws_new:
        for c in cell:
            if c.value in replacements:
                c.value = replacements[c.value]

    wb_new.save(new_file)

import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

def format_currency(amount):
    if amount == int(amount):
        return int(amount)
    return round(amount, 2)

def safe_write(ws, row, col, value):
    """
    Безопасная запись в ячейку с учетом объединенных ячеек
    openpyxl не позволяет писать в объединенные ячейки, нужно писать в левую верхнюю
    """
    # Если значение None или пустая строка, просто возвращаем
    if value is None:
        return
    
    # Пробуем записать напрямую
    try:
        cell = ws.cell(row=row, column=col)
        cell.value = value
        return
    except AttributeError:
        pass
    except Exception:
        pass
    
    # Если не получилось, ищем левую верхнюю ячейку объединенного диапазона
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and \
           merged_range.min_col <= col <= merged_range.max_col:
            # Нашли объединенный диапазон, пишем в левую верхнюю ячейку
            top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            top_left.value = value
            return
    
    # Если не нашли объединение, пробуем еще раз напрямую
    try:
        ws.cell(row=row, column=col).value = value
    except:
        pass


def fill_kudir_template(operations, template_path, output_path, inn, fio, year=2025):
    """Заполнение шаблона КУДиР"""
    wb = load_workbook(template_path)
    
    # Лист 1 (титульный)
    ws1 = wb["Лист1"]
    
    # Год (AD13:AE13) - пишем в левую верхнюю (AD13)
    safe_write(ws1, 13, column_index_from_string('AD'), year)
    
    # ФИО (D15:D16)
    fio_parts = fio.split()
    if len(fio_parts) >= 1:
        safe_write(ws1, 15, 4, fio_parts[0])  # фамилия
    if len(fio_parts) >= 2:
        safe_write(ws1, 16, 4, fio_parts[1] + (" " + fio_parts[2] if len(fio_parts) > 2 else ""))
    
    # ИНН (D20:D21)
    safe_write(ws1, 20, 4, inn)
    
    # Объект налогообложения (D27:D28)
    safe_write(ws1, 27, 4, "Доходы")
    
    # Сортируем операции
    sorted_ops = sorted(operations, key=lambda x: x['date'])
    total_income = sum(op['amount'] for op in sorted_ops)
    
    quarterly_totals = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for op in sorted_ops:
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly_totals[quarter] += op['amount']
    
    # Лист 2 (доходы I-II квартал)
    ws2 = wb["Лист2"]
    
    # Очищаем старые данные в таблице (строки 14-200)
    for row in range(14, 200):
        for col in range(1, 6):
            ws2.cell(row=row, column=col).value = None
    
    # Заполняем операции за I и II кварталы
    row = 14
    for op in sorted_ops:
        quarter = (op['date'].month - 1) // 3 + 1
        if quarter <= 2:
            safe_write(ws2, row, 1, row - 13)
            safe_write(ws2, row, 2, op['document'])
            safe_write(ws2, row, 3, op['purpose'][:150])
            safe_write(ws2, row, 4, format_currency(op['amount']))
            row += 1
    
    # Итоги на Лист2
    for r in range(14, 100):
        cell_val = ws2.cell(row=r, column=1).value
        if cell_val and isinstance(cell_val, str):
            if "Итого за I квартал" in cell_val:
                safe_write(ws2, r, 4, format_currency(quarterly_totals[1]))
            elif "Итого за II квартал" in cell_val:
                safe_write(ws2, r, 4, format_currency(quarterly_totals[2]))
            elif "Итого за полугодие" in cell_val:
                safe_write(ws2, r, 4, format_currency(quarterly_totals[1] + quarterly_totals[2]))
    
    # Лист 3 (доходы III-IV квартал)
    ws3 = wb["Лист3"]
    
    # Очищаем старые данные
    for row in range(14, 200):
        for col in range(1, 6):
            ws3.cell(row=row, column=col).value = None
    
    # Заполняем операции за III и IV кварталы
    row = 14
    for op in sorted_ops:
        quarter = (op['date'].month - 1) // 3 + 1
        if quarter >= 3:
            safe_write(ws3, row, 1, row - 13)
            safe_write(ws3, row, 2, op['document'])
            safe_write(ws3, row, 3, op['purpose'][:150])
            safe_write(ws3, row, 4, format_currency(op['amount']))
            row += 1
    
    # Итоги на Лист3
    for r in range(14, 100):
        cell_val = ws3.cell(row=r, column=1).value
        if cell_val and isinstance(cell_val, str):
            if "Итого за III квартал" in cell_val:
                safe_write(ws3, r, 4, format_currency(quarterly_totals[3]))
            elif "Итого за 9 месяцев" in cell_val:
                safe_write(ws3, r, 4, format_currency(quarterly_totals[1] + quarterly_totals[2] + quarterly_totals[3]))
            elif "Итого за IV квартал" in cell_val:
                safe_write(ws3, r, 4, format_currency(quarterly_totals[4]))
            elif "Итого за год" in cell_val:
                safe_write(ws3, r, 4, format_currency(total_income))
            elif cell_val.strip() == "010":
                safe_write(ws3, r, 4, format_currency(total_income))
            elif cell_val.strip() == "020":
                safe_write(ws3, r, 4, 0)
            elif cell_val.strip() == "040":
                safe_write(ws3, r, 4, format_currency(total_income))
    
    wb.save(output_path)
    return total_income


def fill_declaration_template(operations, ens_data, template_path, output_excel, output_xml, inn, fio, oktmo):
    """Заполнение шаблона декларации"""
    
    # Расчет доходов по кварталам
    quarterly = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for op in operations:
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly[quarter] += op['amount']
    
    total_income = quarterly[1] + quarterly[2] + quarterly[3] + quarterly[4]
    tax_rate = 6
    tax_amount = total_income * tax_rate / 100
    
    paid_in_2025 = any(d.year == 2025 for d in ens_data.get('insurance_paid_dates', []))
    insurance_paid = ens_data.get('insurance_paid', 0)
    
    if paid_in_2025:
        tax_payable = max(0, tax_amount - insurance_paid)
    else:
        tax_payable = tax_amount
    
    # Суммы нарастающим
    cum_income_1 = quarterly[1]
    cum_income_2 = quarterly[1] + quarterly[2]
    cum_income_3 = quarterly[1] + quarterly[2] + quarterly[3]
    cum_income_4 = total_income
    
    cum_tax_1 = cum_income_1 * tax_rate / 100
    cum_tax_2 = cum_income_2 * tax_rate / 100
    cum_tax_3 = cum_income_3 * tax_rate / 100
    cum_tax_4 = tax_amount
    
    wb = load_workbook(template_path)
    ws = wb["стр.1"]
    
    # ИНН (AG7:AG8)
    safe_write(ws, 7, column_index_from_string('AG'), inn)
    
    # Отчетный год (BJ14:BK14)
    safe_write(ws, 14, column_index_from_string('BJ'), 2025)
    
    # ФИО (ищем строку с фамилией)
    for row in range(30, 50):
        cell_val = ws.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str) and "фамилия" in cell_val.lower():
            safe_write(ws, row+1, 1, fio)
            break
    
    # ОКВЭД (ищем строку)
    for row in range(30, 60):
        cell_val = ws.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str) and "ОКВЭД" in cell_val:
            safe_write(ws, row, 2, "47.91")
            break
    
    # Заполнение строк по кодам (колонка C - код, колонка D - значение)
    for row in range(50, 200):
        code_cell = ws.cell(row=row, column=3).value
        if code_cell:
            code = str(code_cell).strip()
            if code == "010":
                safe_write(ws, row, 4, format_currency(cum_income_1))
            elif code == "011":
                safe_write(ws, row, 4, format_currency(cum_income_2))
            elif code == "012":
                safe_write(ws, row, 4, format_currency(cum_income_3))
            elif code == "013":
                safe_write(ws, row, 4, format_currency(cum_income_4))
            elif code == "020":
                safe_write(ws, row, 4, tax_rate)
            elif code == "030":
                safe_write(ws, row, 4, format_currency(cum_tax_1))
            elif code == "031":
                safe_write(ws, row, 4, format_currency(cum_tax_2))
            elif code == "032":
                safe_write(ws, row, 4, format_currency(cum_tax_3))
            elif code == "033":
                safe_write(ws, row, 4, format_currency(cum_tax_4))
            elif code == "040":
                safe_write(ws, row, 4, 0)
            elif code == "041":
                safe_write(ws, row, 4, 0)
            elif code == "042":
                safe_write(ws, row, 4, 0)
            elif code == "043":
                safe_write(ws, row, 4, 0)
            elif code == "050":
                safe_write(ws, row, 4, oktmo)
            elif code == "060":
                safe_write(ws, row, 4, format_currency(tax_payable))
    
    wb.save(output_excel)
    
    # XML
    fio_parts = fio.split()
    last_name = fio_parts[0] if len(fio_parts) > 0 else ""
    first_name = fio_parts[1] if len(fio_parts) > 1 else ""
    patronymic = fio_parts[2] if len(fio_parts) > 2 else ""
    
    xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<Файл xmlns="urn:ФНС-СХД-Декл-УСН-2025-1">
    <Документ>
        <КНД>1152017</КНД>
        <ДатаДок>{datetime.now().strftime('%Y-%m-%d')}</ДатаДок>
        <НомКорр>0</НомКорр>
    </Документ>
    <НалогПериод>
        <НомерПериода>34</НомерПериода>
        <ОтчетныйГод>2025</ОтчетныйГод>
    </НалогПериод>
    <Налогоплательщик>
        <ИНН>{inn}</ИНН>
        <ИП>
            <ФИО>
                <Фамилия>{last_name}</Фамилия>
                <Имя>{first_name}</Имя>
                <Отчество>{patronymic}</Отчество>
            </ФИО>
        </ИП>
    </Налогоплательщик>
    <Показатели>
        <Раздел1_1>
            <ОКТМО>{oktmo}</ОКТМО>
            <СумНал100>{int(tax_payable)}</СумНал100>
        </Раздел1_1>
        <Раздел2_1_1>
            <СумДоход110>{int(cum_income_1)}</СумДоход110>
            <СумДоход111>{int(cum_income_2)}</СумДоход111>
            <СумДоход112>{int(cum_income_3)}</СумДоход112>
            <СумДоход113>{int(cum_income_4)}</СумДоход113>
            <НалСтавка120>{tax_rate}</НалСтавка120>
            <СумИсчисНал130>{int(cum_tax_1)}</СумИсчисНал130>
            <СумИсчисНал131>{int(cum_tax_2)}</СумИсчисНал131>
            <СумИсчисНал132>{int(cum_tax_3)}</СумИсчисНал132>
            <СумИсчисНал133>{int(cum_tax_4)}</СумИсчисНал133>
            <СумУплНал140>0</СумУплНал140>
            <СумУплНал141>0</СумУплНал141>
            <СумУплНал142>0</СумУплНал142>
            <СумУплНал143>0</СумУплНал143>
        </Раздел2_1_1>
    </Показатели>
</Файл>'''
    
    with open(output_xml, 'w', encoding='utf-8') as f:
        f.write(xml)
    
    return tax_payable, total_income


def generate_report(operations, ens_data, output_dir, user_id, kudir_template, decl_template, inn, fio, oktmo):
    """Генерация отчетности"""
    kudir_path = os.path.join(output_dir, f"kudir_{user_id}.xlsx")
    total_income = fill_kudir_template(operations, kudir_template, kudir_path, inn, fio)
    
    decl_excel = os.path.join(output_dir, f"declaration_{user_id}.xlsx")
    decl_xml = os.path.join(output_dir, f"declaration_{user_id}.xml")
    tax_payable, total_income = fill_declaration_template(
        operations, ens_data, decl_template, decl_excel, decl_xml, inn, fio, oktmo
    )
    
    return kudir_path, decl_excel, decl_xml, total_income, tax_payable
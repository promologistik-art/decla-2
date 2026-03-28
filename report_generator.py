import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

IP_INN = "632312967829"
IP_FIO = "Леонтьев Артём Владиславович"
IP_OKTMO = "36701320"

def format_currency(amount):
    if amount == int(amount):
        return int(amount)
    return round(amount, 2)

def safe_write(ws, row, col, value):
    try:
        cell = ws.cell(row=row, column=col)
        cell.value = value
    except AttributeError:
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and \
               merged_range.min_col <= col <= merged_range.max_col:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
        raise


def fill_kudir_template(operations, template_path, output_path, inn=IP_INN, fio=IP_FIO, year=2025):
    wb = load_workbook(template_path)
    ws1 = wb["Лист1"]
    
    safe_write(ws1, 13, column_index_from_string('AD'), year)
    
    fio_parts = fio.split()
    if len(fio_parts) >= 1:
        safe_write(ws1, 15, 4, fio_parts[0])
    if len(fio_parts) >= 2:
        safe_write(ws1, 16, 4, fio_parts[1] + (" " + fio_parts[2] if len(fio_parts) > 2 else ""))
    
    safe_write(ws1, 20, 4, inn)
    safe_write(ws1, 27, 4, "Доходы")
    
    ws2 = wb["Лист2"]
    
    start_row = None
    for row in range(10, 30):
        if ws2.cell(row=row, column=1).value == 1:
            start_row = row
            break
    if start_row is None:
        start_row = 14
    
    sorted_ops = sorted(operations, key=lambda x: x['date'])
    quarterly_totals = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    
    for idx, op in enumerate(sorted_ops, 1):
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly_totals[quarter] += op['amount']
        
        safe_write(ws2, start_row + idx - 1, 1, idx)
        safe_write(ws2, start_row + idx - 1, 2, op['document'])
        safe_write(ws2, start_row + idx - 1, 3, op['purpose'][:150])
        safe_write(ws2, start_row + idx - 1, 4, format_currency(op['amount']))
    
    for row in range(start_row, start_row + len(sorted_ops) + 30):
        cell_val = ws2.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str):
            if "Итого за I квартал" in cell_val:
                safe_write(ws2, row, 4, format_currency(quarterly_totals[1]))
            elif "Итого за II квартал" in cell_val:
                safe_write(ws2, row, 4, format_currency(quarterly_totals[2]))
            elif "Итого за полугодие" in cell_val:
                safe_write(ws2, row, 4, format_currency(quarterly_totals[1] + quarterly_totals[2]))
    
    ws3 = wb["Лист3"]
    total_income = sum(quarterly_totals.values())
    
    for row in range(10, 80):
        cell_val = ws3.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str):
            if "Итого за III квартал" in cell_val:
                safe_write(ws3, row, 4, format_currency(quarterly_totals[3]))
            elif "Итого за 9 месяцев" in cell_val:
                safe_write(ws3, row, 4, format_currency(quarterly_totals[1] + quarterly_totals[2] + quarterly_totals[3]))
            elif "Итого за IV квартал" in cell_val:
                safe_write(ws3, row, 4, format_currency(quarterly_totals[4]))
            elif "Итого за год" in cell_val:
                safe_write(ws3, row, 4, format_currency(total_income))
            elif cell_val.strip() == "010":
                safe_write(ws3, row, 4, format_currency(total_income))
            elif cell_val.strip() == "020":
                safe_write(ws3, row, 4, 0)
            elif cell_val.strip() == "040":
                safe_write(ws3, row, 4, format_currency(total_income))
    
    wb.save(output_path)
    return total_income


def fill_declaration_template(operations, ens_data, template_path, output_excel, output_xml, inn=IP_INN, fio=IP_FIO):
    quarterly = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for op in operations:
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly[quarter] += op['amount']
    
    total_income = sum(quarterly.values())
    tax_rate = 6
    tax_amount = total_income * tax_rate / 100
    
    paid_in_2025 = any(d.year == 2025 for d in ens_data.get('insurance_paid_dates', []))
    insurance_paid = ens_data.get('insurance_paid', 0)
    
    if paid_in_2025:
        tax_payable = max(0, tax_amount - insurance_paid)
        deductible = insurance_paid
    else:
        tax_payable = tax_amount
        deductible = 0
    
    cum_income = {
        1: quarterly[1],
        2: quarterly[1] + quarterly[2],
        3: quarterly[1] + quarterly[2] + quarterly[3],
        4: total_income
    }
    
    cum_tax = {
        1: cum_income[1] * tax_rate / 100,
        2: cum_income[2] * tax_rate / 100,
        3: cum_income[3] * tax_rate / 100,
        4: tax_amount
    }
    
    if paid_in_2025:
        cum_deductible = {
            1: min(cum_tax[1], deductible),
            2: min(cum_tax[2], deductible),
            3: min(cum_tax[3], deductible),
            4: min(cum_tax[4], deductible)
        }
    else:
        cum_deductible = {1: 0, 2: 0, 3: 0, 4: 0}
    
    wb = load_workbook(template_path)
    ws = wb["стр.1"]
    
    safe_write(ws, 7, column_index_from_string('AG'), inn)
    safe_write(ws, 14, column_index_from_string('BJ'), 2025)
    
    for row in range(30, 50):
        cell_val = ws.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str) and "фамилия" in cell_val.lower():
            safe_write(ws, row+1, 1, fio)
            break
    
    for row in range(30, 60):
        cell_val = ws.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str) and "ОКВЭД" in cell_val:
            safe_write(ws, row, 2, "47.91")
            break
    
    for row in range(50, 200):
        code_cell = ws.cell(row=row, column=3).value
        if code_cell:
            code = str(code_cell).strip()
            if code == "010":
                safe_write(ws, row, 4, format_currency(cum_income[1]))
            elif code == "011":
                safe_write(ws, row, 4, format_currency(cum_income[2]))
            elif code == "012":
                safe_write(ws, row, 4, format_currency(cum_income[3]))
            elif code == "013":
                safe_write(ws, row, 4, format_currency(cum_income[4]))
            elif code == "020":
                safe_write(ws, row, 4, tax_rate)
            elif code == "030":
                safe_write(ws, row, 4, format_currency(cum_tax[1]))
            elif code == "031":
                safe_write(ws, row, 4, format_currency(cum_tax[2]))
            elif code == "032":
                safe_write(ws, row, 4, format_currency(cum_tax[3]))
            elif code == "033":
                safe_write(ws, row, 4, format_currency(cum_tax[4]))
            elif code == "040":
                safe_write(ws, row, 4, format_currency(cum_deductible[1]))
            elif code == "041":
                safe_write(ws, row, 4, format_currency(cum_deductible[2]))
            elif code == "042":
                safe_write(ws, row, 4, format_currency(cum_deductible[3]))
            elif code == "043":
                safe_write(ws, row, 4, format_currency(cum_deductible[4]))
            elif code == "050":
                safe_write(ws, row, 4, IP_OKTMO)
            elif code == "060":
                safe_write(ws, row, 4, format_currency(tax_payable))
    
    wb.save(output_excel)
    
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
            <ОКТМО>{IP_OKTMO}</ОКТМО>
            <СумНал100>{int(tax_payable)}</СумНал100>
        </Раздел1_1>
        <Раздел2_1_1>
            <СумДоход110>{int(cum_income[1])}</СумДоход110>
            <СумДоход111>{int(cum_income[2])}</СумДоход111>
            <СумДоход112>{int(cum_income[3])}</СумДоход112>
            <СумДоход113>{int(cum_income[4])}</СумДоход113>
            <НалСтавка120>{tax_rate}</НалСтавка120>
            <СумИсчисНал130>{int(cum_tax[1])}</СумИсчисНал130>
            <СумИсчисНал131>{int(cum_tax[2])}</СумИсчисНал131>
            <СумИсчисНал132>{int(cum_tax[3])}</СумИсчисНал132>
            <СумИсчисНал133>{int(cum_tax[4])}</СумИсчисНал133>
            <СумУплНал140>{int(cum_deductible[1])}</СумУплНал140>
            <СумУплНал141>{int(cum_deductible[2])}</СумУплНал141>
            <СумУплНал142>{int(cum_deductible[3])}</СумУплНал142>
            <СумУплНал143>{int(cum_deductible[4])}</СумУплНал143>
        </Раздел2_1_1>
    </Показатели>
</Файл>'''
    
    with open(output_xml, 'w', encoding='utf-8') as f:
        f.write(xml)
    
    return tax_payable, total_income


# ========== ГЛАВНАЯ ФУНКЦИЯ - 6 АРГУМЕНТОВ ==========
def generate_report(operations, ens_data, output_dir, user_id, kudir_template, decl_template):
    """Генерация отчетности"""
    kudir_path = os.path.join(output_dir, f"kudir_{user_id}.xlsx")
    total_income = fill_kudir_template(operations, kudir_template, kudir_path)
    
    decl_excel = os.path.join(output_dir, f"declaration_{user_id}.xlsx")
    decl_xml = os.path.join(output_dir, f"declaration_{user_id}.xml")
    tax_payable, total_income = fill_declaration_template(
        operations, ens_data, decl_template, decl_excel, decl_xml
    )
    
    return kudir_path, decl_excel, decl_xml, total_income, tax_payable
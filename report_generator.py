import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

IP_OKTMO = "36701320"

def format_currency(amount):
    if amount == int(amount):
        return int(amount)
    return round(amount, 2)

def safe_write(ws, row, col, value):
    """Безопасная запись в ячейку с учетом объединенных ячеек"""
    if value is None:
        return
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            ws.cell(row=merged.min_row, column=merged.min_col).value = value
            return
    ws.cell(row=row, column=col).value = value

def write_inn_digit_by_digit(ws, inn):
    """Записывает ИНН по одной цифре в ячейки A28, C28, E28, G28, I28, K28, M28, O28, Q28, S28, U28, W28"""
    inn_str = str(inn)
    inn_str = ''.join(ch for ch in inn_str if ch.isdigit())
    
    positions = [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23]
    
    for i, digit in enumerate(inn_str):
        if i < len(positions):
            col = positions[i]
            safe_write(ws, 28, col, int(digit))

def fill_kudir_template(operations, template_path, output_path, inn, fio, ip_accounts, year=2025):
    """Заполнение титульного листа шаблона КУДиР"""
    wb = load_workbook(template_path)
    ws1 = wb["Лист1"]
    
    # Год (H15)
    year_last_two = year % 100
    safe_write(ws1, 15, column_index_from_string('H'), year_last_two)
    
    # ФИО (V18)
    safe_write(ws1, 18, column_index_from_string('V'), fio)
    
    # ИНН (A28, C28, E28, G28, I28, K28, M28, O28, Q28, S28, U28, W28)
    write_inn_digit_by_digit(ws1, inn)
    
    # Форма по ОКУД (BB14)
    safe_write(ws1, 14, column_index_from_string('BB'), 1151085)
    
    # Дата заполнения (BB15, BG15, BJ15)
    today = datetime.now()
    safe_write(ws1, 15, column_index_from_string('BB'), today.year)
    safe_write(ws1, 15, column_index_from_string('BG'), today.month)
    safe_write(ws1, 15, column_index_from_string('BJ'), today.day)
    
    # Объект налогообложения (P30)
    safe_write(ws1, 30, column_index_from_string('P'), "Доходы")
    
    # Счета ИП (A38, A40, A42...)
    row = 38
    for acc in ip_accounts:
        account_text = f"{acc['number']} {acc['bank']} БИК {acc['bik']}"
        safe_write(ws1, row, 1, account_text)
        row += 2
    
    wb.save(output_path)
    
    total_income = sum(op['amount'] for op in operations)
    return total_income


def fill_declaration_template(operations, ens_data, template_path, output_excel, output_xml, inn, fio, oktmo, okved, phone):
    """Заполнение шаблона декларации"""
    
    # ========== РАСЧЕТ ДОХОДОВ ПО КВАРТАЛАМ ==========
    quarterly = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for op in operations:
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly[quarter] += op['amount']
    
    total_income = quarterly[1] + quarterly[2] + quarterly[3] + quarterly[4]
    tax_rate = 6
    tax_amount = total_income * tax_rate / 100
    
    # ========== АВАНСОВЫЕ ПЛАТЕЖИ ИЗ ЕНС ==========
    usn_payments = ens_data.get('usn_payments', [])
    advance_payments = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    
    for payment in usn_payments:
        # Определяем период аванса по дате
        if payment['date']:
            month = payment['date'].month
            if month <= 3:
                advance_payments[1] += payment['amount']
            elif month <= 6:
                advance_payments[2] += payment['amount']
            elif month <= 9:
                advance_payments[3] += payment['amount']
            else:
                advance_payments[4] += payment['amount']
    
    # ========== СУММЫ НАРАСТАЮЩИМ ИТОГОМ ==========
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
    
    # Вычеты по взносам (только уплаченные в 2025)
    paid_in_2025 = any(d.year == 2025 for d in ens_data.get('insurance_paid_dates', []))
    insurance_paid = ens_data.get('insurance_paid', 0)
    
    if paid_in_2025:
        cum_deductible = {
            1: min(cum_tax[1], insurance_paid),
            2: min(cum_tax[2], insurance_paid),
            3: min(cum_tax[3], insurance_paid),
            4: min(cum_tax[4], insurance_paid)
        }
    else:
        cum_deductible = {1: 0, 2: 0, 3: 0, 4: 0}
    
    # Налог к уплате за год (с учетом авансов)
    tax_payable = max(0, cum_tax[4] - cum_deductible[4] - advance_payments[1] - advance_payments[2] - advance_payments[3])
    
    # ========== ЗАПОЛНЕНИЕ EXCEL ==========
    wb = load_workbook(template_path)
    ws = wb["стр.1"]
    
    # ИНН (AG7)
    safe_write(ws, 7, column_index_from_string('AG'), inn)
    
    # Отчетный год (BJ14)
    safe_write(ws, 14, column_index_from_string('BJ'), 2025)
    
    # Контактный телефон
    if phone:
        for row in range(30, 60):
            cell_val = ws.cell(row=row, column=1).value
            if cell_val and isinstance(cell_val, str) and "телефон" in cell_val.lower():
                safe_write(ws, row, 2, phone)
                break
    
    # ФИО (ищем строку с "фамилия, имя, отчество")
    for row in range(30, 60):
        cell_val = ws.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str) and "фамилия" in cell_val.lower():
            safe_write(ws, row+1, 3, fio)
            break
    
    # ОКВЭД
    if okved:
        for row in range(30, 60):
            cell_val = ws.cell(row=row, column=1).value
            if cell_val and isinstance(cell_val, str) and "ОКВЭД" in cell_val:
                safe_write(ws, row, 2, okved)
                break
    
    # Заполнение строк по кодам
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
                safe_write(ws, row, 4, oktmo)
            elif code == "060":
                safe_write(ws, row, 4, format_currency(tax_payable))
    
    wb.save(output_excel)
    
    # ========== XML ==========
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
            <СумАван010>{int(advance_payments[1])}</СумАван010>
            <СумАван020>{int(advance_payments[2])}</СумАван020>
            <СумАван040>{int(advance_payments[3])}</СумАван040>
            <СумАван070>0</СумАван070>
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


def generate_report(operations, ens_data, output_dir, user_id, kudir_template, decl_template, inn, fio, oktmo, ip_accounts, okved="", phone=""):
    """Генерация отчетности"""
    kudir_path = os.path.join(output_dir, f"kudir_{user_id}.xlsx")
    total_income = fill_kudir_template(operations, kudir_template, kudir_path, inn, fio, ip_accounts)
    
    decl_excel = os.path.join(output_dir, f"declaration_{user_id}.xlsx")
    decl_xml = os.path.join(output_dir, f"declaration_{user_id}.xml")
    tax_payable, total_income = fill_declaration_template(
        operations, ens_data, decl_template, decl_excel, decl_xml, inn, fio, oktmo, okved, phone
    )
    
    return kudir_path, decl_excel, decl_xml, total_income, tax_payable
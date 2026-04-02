import os
import warnings
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def format_currency(amount):
    if amount == int(amount):
        return int(amount)
    return round(amount, 2)

def safe_write(ws, row, col, value):
    if value is None:
        return
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            ws.cell(row=merged.min_row, column=merged.min_col).value = value
            return
    ws.cell(row=row, column=col).value = value


# ========== КУДиР ==========

def write_inn_digit_by_digit_kudir(ws, inn):
    inn_str = ''.join(ch for ch in str(inn) if ch.isdigit())
    positions = [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23]
    for i, digit in enumerate(inn_str):
        if i < len(positions):
            safe_write(ws, 28, positions[i], int(digit))

def fill_kudir_template(operations, template_path, output_path, inn, fio, ip_accounts, year=2025):
    wb = load_workbook(template_path)
    ws1 = wb["Лист1"]
    
    safe_write(ws1, 15, column_index_from_string('H'), year % 100)
    safe_write(ws1, 18, column_index_from_string('V'), fio)
    write_inn_digit_by_digit_kudir(ws1, inn)
    safe_write(ws1, 14, column_index_from_string('BB'), 1151085)
    
    today = datetime.now()
    safe_write(ws1, 15, column_index_from_string('BB'), today.year)
    safe_write(ws1, 15, column_index_from_string('BG'), today.month)
    safe_write(ws1, 15, column_index_from_string('BJ'), today.day)
    safe_write(ws1, 30, column_index_from_string('P'), "Доходы")
    
    row = 38
    for acc in ip_accounts:
        safe_write(ws1, row, 1, f"{acc['number']} {acc['bank']} БИК {acc['bik']}")
        row += 2
    
    wb.save(output_path)
    return sum(op['amount'] for op in operations)


# ========== ДЕКЛАРАЦИЯ ==========

def write_inn_digit_by_digit_declaration(ws, inn):
    inn_str = ''.join(ch for ch in str(inn) if ch.isdigit())
    # Колонки для ИНН в строке 2 на листе "Титул"
    columns = [40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84]
    for i, digit in enumerate(inn_str):
        if i < len(columns):
            safe_write(ws, 2, columns[i], int(digit))

def write_kpp_digit_by_digit(ws, kpp):
    kpp_str = ''.join(ch for ch in str(kpp) if ch.isdigit())
    columns = [40, 44, 48, 52, 56, 60, 64, 68, 72]
    for i, digit in enumerate(kpp_str):
        if i < len(columns):
            safe_write(ws, 4, columns[i], int(digit))

def write_okved_digit_by_digit(ws, okved):
    okved_str = ''.join(ch for ch in str(okved) if ch.isdigit())
    # ОКВЭД в строке 27 на листе "Титул"
    columns = [74, 78, 86, 90, 98, 102]
    for i, digit in enumerate(okved_str):
        if i < len(columns):
            safe_write(ws, 27, columns[i], int(digit))

def write_year_digits(ws, year):
    year_str = str(year)
    # Год в строке 14 на листе "Титул", колонки 114, 118, 122, 126
    columns = [114, 118, 122, 126]
    for i, digit in enumerate(year_str):
        if i < len(columns):
            safe_write(ws, 14, columns[i], int(digit))

def fill_declaration_template(operations, ens_data, template_path, output_excel, output_xml, inn, fio, oktmo, okved, phone):
    wb = load_workbook(template_path)
    
    # Проверяем наличие листа "Титул"
    if "Титул" not in wb.sheetnames:
        raise Exception(f"Лист 'Титул' не найден в шаблоне. Доступные листы: {wb.sheetnames}")
    
    ws = wb["Титул"]
    
    # ИНН
    write_inn_digit_by_digit_declaration(ws, inn)
    
    # Год
    write_year_digits(ws, 2025)
    
    # Номер корректировки (0 - первичная) - строка 14, колонка 18 (R)
    safe_write(ws, 14, 18, 0)
    
    # Телефон (строка 43, колонка AZ = 52)
    if phone:
        phone_digits = ''.join(ch for ch in phone if ch.isdigit())
        # Записываем телефон цифра за цифрой
        for i, digit in enumerate(phone_digits[:11]):
            safe_write(ws, 43, 52 + i, int(digit))
    
    # ФИО в строке 20 (колонка C = 3)
    safe_write(ws, 20, 3, fio)
    
    # ФИО в строке 50 (колонка U = 21)
    safe_write(ws, 50, 21, fio)
    
    # ОКВЭД
    if okved:
        write_okved_digit_by_digit(ws, okved)
    
    # Расчет доходов по кварталам
    quarterly = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for op in operations:
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly[quarter] += op['amount']
    
    total_income = sum(quarterly.values())
    tax_rate = 6
    tax_amount = total_income * tax_rate / 100
    
    # Авансовые платежи из ЕНС
    usn_payments = ens_data.get('usn_payments', [])
    advance_payments = {1: 0.0, 2: 0.0, 3: 0.0}
    for payment in usn_payments:
        if payment['date']:
            month = payment['date'].month
            if month <= 3:
                advance_payments[1] += payment['amount']
            elif month <= 6:
                advance_payments[2] += payment['amount']
            elif month <= 9:
                advance_payments[3] += payment['amount']
    
    # Вычет по взносам (только уплаченные в 2025)
    paid_in_2025 = any(d.year == 2025 for d in ens_data.get('insurance_paid_dates', []))
    insurance_paid = ens_data.get('insurance_paid', 0) if paid_in_2025 else 0
    
    # Накопленные доходы
    cum_income = {
        1: quarterly[1],
        2: quarterly[1] + quarterly[2],
        3: quarterly[1] + quarterly[2] + quarterly[3],
        4: total_income
    }
    
    # Накопленный налог
    cum_tax = {i: cum_income[i] * tax_rate / 100 for i in range(1, 5)}
    
    # Накопленный вычет (не более налога)
    cum_deductible = {i: min(cum_tax[i], insurance_paid) for i in range(1, 5)} if paid_in_2025 else {i: 0 for i in range(1, 5)}
    
    # Налог к уплате
    tax_payable = max(0, cum_tax[4] - cum_deductible[4] - advance_payments[1] - advance_payments[2] - advance_payments[3])
    
    # Заполнение раздела 2.1.1 (лист "Раздел 2.1.1")
    if "Раздел 2.1.1" not in wb.sheetnames:
        raise Exception(f"Лист 'Раздел 2.1.1' не найден. Доступные листы: {wb.sheetnames}")
    
    ws21 = wb["Раздел 2.1.1"]
    
    # Доходы (строки 110-113)
    safe_write(ws21, 34, 39, format_currency(cum_income[1]))   # стр 110
    safe_write(ws21, 35, 39, format_currency(cum_income[2]))   # стр 111
    safe_write(ws21, 36, 39, format_currency(cum_income[3]))   # стр 112
    safe_write(ws21, 37, 39, format_currency(cum_income[4]))   # стр 113
    
    # Ставка (строки 120-123)
    safe_write(ws21, 41, 39, tax_rate)   # стр 120
    safe_write(ws21, 42, 39, tax_rate)   # стр 121
    safe_write(ws21, 43, 39, tax_rate)   # стр 122
    safe_write(ws21, 44, 39, tax_rate)   # стр 123
    
    # Исчисленный налог (строки 130-133)
    safe_write(ws21, 50, 39, format_currency(cum_tax[1]))   # стр 130
    safe_write(ws21, 51, 39, format_currency(cum_tax[2]))   # стр 131
    safe_write(ws21, 52, 39, format_currency(cum_tax[3]))   # стр 132
    safe_write(ws21, 53, 39, format_currency(cum_tax[4]))   # стр 133
    
    # Вычет по взносам (лист "Раздел 2.1.1 (продолжение)")
    if "Раздел 2.1.1 (продолжение)" not in wb.sheetnames:
        raise Exception(f"Лист 'Раздел 2.1.1 (продолжение)' не найден. Доступные листы: {wb.sheetnames}")
    
    ws21_cont = wb["Раздел 2.1.1 (продолжение)"]
    safe_write(ws21_cont, 12, 39, format_currency(cum_deductible[1]))   # стр 140
    safe_write(ws21_cont, 14, 39, format_currency(cum_deductible[2]))   # стр 141
    safe_write(ws21_cont, 16, 39, format_currency(cum_deductible[3]))   # стр 142
    safe_write(ws21_cont, 18, 39, format_currency(cum_deductible[4]))   # стр 143
    
    # Заполнение раздела 1.1 (лист "Раздел 1.1")
    if "Раздел 1.1" not in wb.sheetnames:
        raise Exception(f"Лист 'Раздел 1.1' не найден. Доступные листы: {wb.sheetnames}")
    
    ws11 = wb["Раздел 1.1"]
    
    # ОКТМО (строка 010)
    safe_write(ws11, 22, 39, oktmo)
    
    # Авансовые платежи
    safe_write(ws11, 28, 39, format_currency(advance_payments[1]))   # стр 020 (28 апреля)
    safe_write(ws11, 38, 39, format_currency(advance_payments[2]))   # стр 040 (28 июля)
    safe_write(ws11, 54, 39, format_currency(advance_payments[3]))   # стр 070 (28 октября)
    
    # Налог к уплате за год (стр 100)
    safe_write(ws11, 70, 39, format_currency(tax_payable))
    
    # Сохраняем Excel
    wb.save(output_excel)
    
    # Генерируем XML
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
    kudir_path = os.path.join(output_dir, f"kudir_{user_id}.xlsx")
    total_income = fill_kudir_template(operations, kudir_template, kudir_path, inn, fio, ip_accounts)
    
    decl_excel = os.path.join(output_dir, f"declaration_{user_id}.xlsx")
    decl_xml = os.path.join(output_dir, f"declaration_{user_id}.xml")
    tax_payable, total_income = fill_declaration_template(
        operations, ens_data, decl_template, decl_excel, decl_xml, inn, fio, oktmo, okved, phone
    )
    
    return kudir_path, decl_excel, decl_xml, total_income, tax_payable
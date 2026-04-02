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
    """Заполнение ИНН: колонки Y1, AA1, AC1, AE1, AG1, AI1, AK1, AM1, AO1, AQ1, AS1, AU1"""
    inn_str = ''.join(ch for ch in str(inn) if ch.isdigit())
    columns = [25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45, 47]
    for i, digit in enumerate(inn_str):
        if i < len(columns):
            safe_write(ws, 1, columns[i], int(digit))

def write_kpp_digit_by_digit_declaration(ws, kpp):
    """Заполнение КПП для организации (для ИП не используется)"""
    kpp_str = ''.join(ch for ch in str(kpp) if ch.isdigit())
    columns = [25, 27, 29, 31, 33, 35, 37, 39, 41]
    for i, digit in enumerate(kpp_str):
        if i < len(columns):
            safe_write(ws, 4, columns[i], int(digit))

def write_tax_office_code(ws, inn):
    """Заполнение кода налогового органа (первые 4 цифры ИНН) в ячейки AA13, AC13, AE13, AG13"""
    inn_str = ''.join(ch for ch in str(inn) if ch.isdigit())
    tax_code = inn_str[:4]
    columns = [27, 29, 31, 33]  # AA, AC, AE, AG
    for i, digit in enumerate(tax_code):
        if i < len(columns):
            safe_write(ws, 13, columns[i], int(digit))

def write_place_of_registration_code(ws):
    """Код по месту учета для ИП: 120 в ячейки BW13, BY13, CA13"""
    # 1, 2, 0
    safe_write(ws, 13, 75, 1)  # BW
    safe_write(ws, 13, 77, 2)  # BY
    safe_write(ws, 13, 79, 0)  # CA

def write_correction_number(ws):
    """Номер корректировки в ячейку S11"""
    safe_write(ws, 11, 19, 0)  # S = 19

def write_tax_period_code(ws):
    """Налоговый период (код) 34 в ячейки BA11, BC11"""
    safe_write(ws, 11, 53, 3)  # BA = 53
    safe_write(ws, 11, 55, 4)  # BC = 55

def write_report_year(ws, year):
    """Отчетный год в ячейки BZ11, CB11, CD11, CF11 (2025)"""
    year_str = str(year)
    columns = [78, 80, 82, 84]  # BZ, CB, CD, CF
    for i, digit in enumerate(year_str):
        if i < len(columns):
            safe_write(ws, 11, columns[i], int(digit))

def write_legal_name_by_letters(ws, name):
    """Заполнение полного названия юрлица по буквам с A15 через одну колонку до CA15, затем перенос на A17"""
    name_clean = ''.join(ch for ch in name if ch.isalpha() or ch == ' ')
    row = 15
    col = 1  # A
    for char in name_clean:
        if char == ' ':
            char = '_'
        if col > 79:  # CA = 79
            row = 17
            col = 1
        safe_write(ws, row, col, char.upper())
        col += 2

def write_phone_by_letters(ws, phone):
    """Заполнение номера телефона с U27 через одну колонку"""
    phone_digits = ''.join(ch for ch in str(phone) if ch.isdigit())
    col = 21  # U
    for digit in phone_digits[:11]:
        safe_write(ws, 27, col, int(digit))
        col += 2

def write_last_name_by_letters(ws, last_name):
    """Заполнение фамилии с B43 через одну колонку"""
    col = 2  # B
    for char in last_name.upper():
        safe_write(ws, 43, col, char)
        col += 2

def write_first_name_by_letters(ws, first_name):
    """Заполнение имени с B45 через одну колонку"""
    col = 2  # B
    for char in first_name.upper():
        safe_write(ws, 45, col, char)
        col += 2

def write_patronymic_by_letters(ws, patronymic):
    """Заполнение отчества с B47 через одну колонку"""
    col = 2  # B
    for char in patronymic.upper():
        safe_write(ws, 47, col, char)
        col += 2

def write_signature_last_name(ws, last_name):
    """Фамилия подписанта в ячейку H50"""
    safe_write(ws, 50, 8, last_name.upper())  # H = 8

def write_signature_date(ws):
    """Дата подписи: день в V50, X50, месяц в AB50, AD50, год в AH50, AJ50, AL50, AN50"""
    today = datetime.now()
    day = str(today.day).zfill(2)
    month = str(today.month).zfill(2)
    year = str(today.year)
    
    # День: V50 (22), X50 (24)
    safe_write(ws, 50, 22, int(day[0]))
    safe_write(ws, 50, 24, int(day[1]))
    
    # Месяц: AB50 (28), AD50 (30)
    safe_write(ws, 50, 28, int(month[0]))
    safe_write(ws, 50, 30, int(month[1]))
    
    # Год: AH50 (34), AJ50 (36), AL50 (38), AN50 (40)
    safe_write(ws, 50, 34, int(year[0]))
    safe_write(ws, 50, 36, int(year[1]))
    safe_write(ws, 50, 38, int(year[2]))
    safe_write(ws, 50, 40, int(year[3]))

def fill_declaration_template(operations, ens_data, template_path, output_excel, output_xml, inn, fio, oktmo, okved, phone):
    wb = load_workbook(template_path)
    
    if "Титул" not in wb.sheetnames:
        raise Exception(f"Лист 'Титул' не найден. Доступные листы: {wb.sheetnames}")
    
    ws = wb["Титул"]
    
    # 1. ИНН
    write_inn_digit_by_digit_declaration(ws, inn)
    
    # 2. КПП - для ИП не заполняем (оставляем пустым)
    
    # 3. Код налогового органа (первые 4 цифры ИНН)
    write_tax_office_code(ws, inn)
    
    # 4. Код по месту учета 120
    write_place_of_registration_code(ws)
    
    # 5. Номер корректировки 0 в S11
    write_correction_number(ws)
    
    # 6. Налоговый период 34 в BA11, BC11
    write_tax_period_code(ws)
    
    # 7. Отчетный год 2025
    write_report_year(ws, 2025)
    
    # 8. Название юрлица по буквам
    write_legal_name_by_letters(ws, f"ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ {fio}")
    
    # 9. Телефон
    if phone:
        write_phone_by_letters(ws, phone)
    
    # 10. Объект налогообложения (1 = доходы) в AJ20
    safe_write(ws, 20, 36, 1)  # AJ = 36
    
    # 11. Разбор ФИО
    fio_parts = fio.split()
    last_name = fio_parts[0] if len(fio_parts) > 0 else ""
    first_name = fio_parts[1] if len(fio_parts) > 1 else ""
    patronymic = fio_parts[2] if len(fio_parts) > 2 else ""
    
    # 12. Фамилия по буквам
    if last_name:
        write_last_name_by_letters(ws, last_name)
    
    # 13. Имя по буквам
    if first_name:
        write_first_name_by_letters(ws, first_name)
    
    # 14. Отчество по буквам
    if patronymic:
        write_patronymic_by_letters(ws, patronymic)
    
    # 15. Фамилия подписанта в H50
    write_signature_last_name(ws, last_name)
    
    # 16. Дата подписи
    write_signature_date(ws)
    
    # Расчет доходов по кварталам
    quarterly = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for op in operations:
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly[quarter] += op['amount']
    
    total_income = sum(quarterly.values())
    tax_rate = 6
    
    cum_income = {
        1: quarterly[1],
        2: quarterly[1] + quarterly[2],
        3: quarterly[1] + quarterly[2] + quarterly[3],
        4: total_income
    }
    
    cum_tax = {i: cum_income[i] * tax_rate / 100 for i in range(1, 5)}
    
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
    
    paid_in_2025 = any(d.year == 2025 for d in ens_data.get('insurance_paid_dates', []))
    insurance_paid = ens_data.get('insurance_paid', 0) if paid_in_2025 else 0
    cum_deductible = {i: min(cum_tax[i], insurance_paid) for i in range(1, 5)} if paid_in_2025 else {i: 0 for i in range(1, 5)}
    
    tax_payable = max(0, cum_tax[4] - cum_deductible[4] - advance_payments[1] - advance_payments[2] - advance_payments[3])
    
    # Заполнение раздела 2.1.1
    if "Раздел 2.1.1" in wb.sheetnames:
        ws21 = wb["Раздел 2.1.1"]
        safe_write(ws21, 34, 39, format_currency(cum_income[1]))
        safe_write(ws21, 35, 39, format_currency(cum_income[2]))
        safe_write(ws21, 36, 39, format_currency(cum_income[3]))
        safe_write(ws21, 37, 39, format_currency(cum_income[4]))
        safe_write(ws21, 41, 39, tax_rate)
        safe_write(ws21, 42, 39, tax_rate)
        safe_write(ws21, 43, 39, tax_rate)
        safe_write(ws21, 44, 39, tax_rate)
        safe_write(ws21, 50, 39, format_currency(cum_tax[1]))
        safe_write(ws21, 51, 39, format_currency(cum_tax[2]))
        safe_write(ws21, 52, 39, format_currency(cum_tax[3]))
        safe_write(ws21, 53, 39, format_currency(cum_tax[4]))
    
    if "Раздел 2.1.1 (продолжение)" in wb.sheetnames:
        ws21_cont = wb["Раздел 2.1.1 (продолжение)"]
        safe_write(ws21_cont, 12, 39, format_currency(cum_deductible[1]))
        safe_write(ws21_cont, 14, 39, format_currency(cum_deductible[2]))
        safe_write(ws21_cont, 16, 39, format_currency(cum_deductible[3]))
        safe_write(ws21_cont, 18, 39, format_currency(cum_deductible[4]))
    
    if "Раздел 1.1" in wb.sheetnames:
        ws11 = wb["Раздел 1.1"]
        safe_write(ws11, 22, 39, oktmo)
        safe_write(ws11, 28, 39, format_currency(advance_payments[1]))
        safe_write(ws11, 38, 39, format_currency(advance_payments[2]))
        safe_write(ws11, 54, 39, format_currency(advance_payments[3]))
        safe_write(ws11, 70, 39, format_currency(tax_payable))
    
    wb.save(output_excel)
    
    # XML
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
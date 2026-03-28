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
    """Безопасная запись в ячейку с учетом объединенных ячеек"""
    if value is None:
        return
    # Проверяем, не входит ли ячейка в объединенный диапазон
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            # Записываем в левую верхнюю ячейку
            ws.cell(row=merged.min_row, column=merged.min_col).value = value
            return
    # Если не объединена, пишем напрямую
    ws.cell(row=row, column=col).value = value


def write_inn_digit_by_digit(ws, start_row, start_col, inn):
    """Записывает ИНН по одной цифре в ячейку"""
    inn_str = str(inn)
    for i, digit in enumerate(inn_str):
        if digit.isdigit():
            safe_write(ws, start_row, start_col + i, int(digit))


def fill_kudir_template(operations, template_path, output_path, inn, fio, ip_accounts, year=2025):
    """
    Заполнение титульного листа шаблона КУДиР
    
    ip_accounts: список словарей с ключами 'number', 'bank', 'bik'
    """
    wb = load_workbook(template_path)
    
    # ========== ЛИСТ 1 (ТИТУЛЬНЫЙ) ==========
    ws1 = wb["Лист1"]
    
    # 1. Год (H15) — последние 2 цифры, "20" уже есть в шаблоне
    year_last_two = year % 100
    safe_write(ws1, 15, column_index_from_string('H'), year_last_two)
    
    # 2. ФИО (V18)
    safe_write(ws1, 18, column_index_from_string('V'), fio)
    
    # 3. ИНН (A28:AA28) — по одной цифре в ячейку
    write_inn_digit_by_digit(ws1, 28, 1, inn)
    
    # 4. Форма по ОКУД (BB14)
    safe_write(ws1, 14, column_index_from_string('BB'), 1151085)
    
    # 5. Дата заполнения (BB15, BG15, BJ15)
    today = datetime.now()
    safe_write(ws1, 15, column_index_from_string('BB'), today.year)
    safe_write(ws1, 15, column_index_from_string('BG'), today.month)
    safe_write(ws1, 15, column_index_from_string('BJ'), today.day)
    
    # 6. Объект налогообложения (P30)
    safe_write(ws1, 30, column_index_from_string('P'), "Доходы")
    
    # 7. Счета ИП (A38, A40, A42...)
    row = 38
    for acc in ip_accounts:
        account_text = f"{acc['number']} {acc['bank']} БИК {acc['bik']}"
        safe_write(ws1, row, 1, account_text)  # col=1 = колонка A
        row += 2
    
    # Сохраняем (пока только титульный лист)
    wb.save(output_path)
    
    # Возвращаем общую сумму доходов
    total_income = sum(op['amount'] for op in operations)
    return total_income


def fill_declaration_template(operations, ens_data, template_path, output_excel, output_xml, inn, fio, oktmo):
    """Заполнение шаблона декларации (пока заглушка)"""
    # Пока просто копируем шаблон
    wb = load_workbook(template_path)
    wb.save(output_excel)
    
    # XML заглушка
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
            <СумНал100>0</СумНал100>
        </Раздел1_1>
        <Раздел2_1_1>
            <СумДоход110>0</СумДоход110>
            <СумДоход111>0</СумДоход111>
            <СумДоход112>0</СумДоход112>
            <СумДоход113>0</СумДоход113>
            <НалСтавка120>6</НалСтавка120>
            <СумИсчисНал130>0</СумИсчисНал130>
            <СумИсчисНал131>0</СумИсчисНал131>
            <СумИсчисНал132>0</СумИсчисНал132>
            <СумИсчисНал133>0</СумИсчисНал133>
            <СумУплНал140>0</СумУплНал140>
            <СумУплНал141>0</СумУплНал141>
            <СумУплНал142>0</СумУплНал142>
            <СумУплНал143>0</СумУплНал143>
        </Раздел2_1_1>
    </Показатели>
</Файл>'''
    
    with open(output_xml, 'w', encoding='utf-8') as f:
        f.write(xml)
    
    return 0, 0  # tax_payable, total_income


def generate_report(operations, ens_data, output_dir, user_id, kudir_template, decl_template, inn, fio, oktmo, ip_accounts):
    """Генерация отчетности"""
    kudir_path = os.path.join(output_dir, f"kudir_{user_id}.xlsx")
    total_income = fill_kudir_template(operations, kudir_template, kudir_path, inn, fio, ip_accounts)
    
    decl_excel = os.path.join(output_dir, f"declaration_{user_id}.xlsx")
    decl_xml = os.path.join(output_dir, f"declaration_{user_id}.xml")
    tax_payable, total_income = fill_declaration_template(
        operations, ens_data, decl_template, decl_excel, decl_xml, inn, fio, oktmo
    )
    
    return kudir_path, decl_excel, decl_xml, total_income, tax_payable
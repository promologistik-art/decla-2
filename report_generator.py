import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

IP_INN = "632312967829"
IP_FIO = "Леонтьев Артём Владиславович"
IP_OKTMO = "36701320"

def format_currency(amount):
    """Безопасное форматирование числа"""
    try:
        # Если пришла строка, преобразуем в число
        if isinstance(amount, str):
            amount = float(amount.replace(" ", "").replace(",", "."))
        # Округляем до 2 знаков, если не целое
        if amount == int(amount):
            return int(amount)
        return round(amount, 2)
    except:
        return 0

def generate_report(operations_data, ens_data, output_dir, user_id):
    """
    Генерация КУДиР и декларации
    operations_data: список словарей с ключами date, amount, purpose
    """
    
    # Нормализуем входные данные - преобразуем в список словарей
    all_ops = []
    
    if isinstance(operations_data, list):
        for item in operations_data:
            if isinstance(item, dict):
                # Проверяем наличие всех нужных ключей
                if 'date' in item and 'amount' in item:
                    all_ops.append({
                        'date': item['date'],
                        'amount': float(item['amount']),
                        'purpose': str(item.get('purpose', ''))
                    })
            elif isinstance(item, list):
                for subitem in item:
                    if isinstance(subitem, dict) and 'date' in subitem and 'amount' in subitem:
                        all_ops.append({
                            'date': subitem['date'],
                            'amount': float(subitem['amount']),
                            'purpose': str(subitem.get('purpose', ''))
                        })
    
    if not all_ops:
        raise Exception("Нет данных для формирования отчетности")
    
    # Сортируем по дате
    all_ops.sort(key=lambda x: x['date'])
    
    # Расчет доходов по кварталам
    quarterly = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for op in all_ops:
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly[quarter] += op['amount']
    
    total_income = quarterly[1] + quarterly[2] + quarterly[3] + quarterly[4]
    tax_rate = 6
    tax_amount = total_income * tax_rate / 100
    
    # Проверка уплаты взносов в 2025
    paid_dates = ens_data.get('insurance_paid_dates', [])
    paid_in_2025 = False
    for d in paid_dates:
        if d and hasattr(d, 'year') and d.year == 2025:
            paid_in_2025 = True
            break
    
    insurance_paid = ens_data.get('insurance_paid', 0)
    if isinstance(insurance_paid, str):
        insurance_paid = float(insurance_paid.replace(" ", "").replace(",", "."))
    
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
    
    # ========== КУДиР ==========
    kudir_path = os.path.join(output_dir, f"kudir_{user_id}.xlsx")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "КУДиР"
    
    # Заголовок
    ws['A1'] = "Книга учета доходов и расходов"
    ws['A2'] = f"ИП {IP_FIO}"
    ws['A3'] = f"ИНН {IP_INN}"
    ws['A4'] = "за 2025 год"
    ws['A5'] = "Объект налогообложения: Доходы"
    
    # Таблица
    headers = ['№ п/п', 'Дата', 'Содержание операции', 'Сумма дохода']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    total = 0
    for idx, op in enumerate(all_ops, 1):
        ws.cell(row=7 + idx, column=1, value=idx)
        ws.cell(row=7 + idx, column=2, value=op['date'].strftime('%d.%m.%Y'))
        purpose = op['purpose'][:150] if len(op['purpose']) > 150 else op['purpose']
        ws.cell(row=7 + idx, column=3, value=purpose)
        amount_val = format_currency(op['amount'])
        ws.cell(row=7 + idx, column=4, value=amount_val)
        total += op['amount']
    
    # Итог
    total_row = 7 + len(all_ops) + 1
    ws.cell(row=total_row, column=3, value="ИТОГО:")
    ws.cell(row=total_row, column=3).font = Font(bold=True)
    ws.cell(row=total_row, column=4, value=format_currency(total))
    
    for col in range(1, 5):
        ws.column_dimensions[chr(64 + col)].width = 20
    
    wb.save(kudir_path)
    
    # ========== Декларация Excel ==========
    decl_excel = os.path.join(output_dir, f"declaration_{user_id}.xlsx")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Декларация УСН"
    
    ws['A1'] = "Налоговая декларация по УСН"
    ws['A2'] = f"ИП {IP_FIO}"
    ws['A3'] = f"ИНН {IP_INN}"
    ws['A4'] = "за 2025 год"
    
    # Раздел 2.1.1
    ws['A6'] = "Раздел 2.1.1. Доходы"
    ws['A6'].font = Font(bold=True)
    
    data_211 = [
        ("Доход за 1 квартал", "110", format_currency(cum_income_1)),
        ("Доход за полугодие", "111", format_currency(cum_income_2)),
        ("Доход за 9 месяцев", "112", format_currency(cum_income_3)),
        ("Доход за год", "113", format_currency(cum_income_4)),
        ("Налоговая ставка (%)", "120", tax_rate),
        ("Сумма налога за 1 квартал", "130", format_currency(cum_tax_1)),
        ("Сумма налога за полугодие", "131", format_currency(cum_tax_2)),
        ("Сумма налога за 9 месяцев", "132", format_currency(cum_tax_3)),
        ("Сумма налога за год", "133", format_currency(cum_tax_4)),
        ("Сумма страховых взносов за год", "143", 0),
    ]
    
    for idx, (name, code, val) in enumerate(data_211, 8):
        ws.cell(row=idx, column=1, value=name)
        ws.cell(row=idx, column=2, value=code)
        ws.cell(row=idx, column=3, value=val)
    
    # Раздел 1.1
    row_start = 8 + len(data_211) + 2
    ws.cell(row=row_start, column=1, value="Раздел 1.1. Сумма налога к уплате")
    ws.cell(row=row_start, column=1).font = Font(bold=True)
    
    data_11 = [
        ("Код ОКТМО", "010", IP_OKTMO),
        ("Налог к уплате за год", "100", format_currency(tax_payable)),
    ]
    
    for idx, (name, code, val) in enumerate(data_11, row_start + 2):
        ws.cell(row=idx, column=1, value=name)
        ws.cell(row=idx, column=2, value=code)
        ws.cell(row=idx, column=3, value=val)
    
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 20
    
    wb.save(decl_excel)
    
    # ========== XML ==========
    decl_xml = os.path.join(output_dir, f"declaration_{user_id}.xml")
    
    fio_parts = IP_FIO.split()
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
        <ИНН>{IP_INN}</ИНН>
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
    
    with open(decl_xml, 'w', encoding='utf-8') as f:
        f.write(xml)
    
    return kudir_path, decl_excel, decl_xml, total_income, tax_payable
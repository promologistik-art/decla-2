import pandas as pd
from datetime import datetime

def safe_float(val):
    try:
        if pd.isna(val):
            return 0.0
        if isinstance(val, str):
            cleaned = val.replace(" ", "").replace(",", ".")
            return float(cleaned)
        return float(val)
    except:
        return 0.0

def parse_date(val):
    if isinstance(val, datetime):
        return val
    if isinstance(val, pd.Timestamp):
        return val.to_pydatetime()
    if isinstance(val, str):
        val = val.strip()
        formats = ["%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M"]
        for fmt in formats:
            try:
                return datetime.strptime(val, fmt)
            except:
                continue
    return None

def extract_ip_data(df):
    """Извлекает ИНН и ФИО из выписки"""
    inn = ""
    fio = ""
    
    for idx, row in df.iterrows():
        for col in range(len(row)):
            val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
            val_lower = val.lower()
            
            # Ищем ФИО (ВБ Банк)
            if "индивидуальный предприниматель" in val_lower:
                # Извлекаем ФИО
                fio = val.replace("Индивидуальный предприниматель", "").replace("ИП", "").strip()
                # Проверяем соседние колонки
                for c in range(max(0, col-2), min(len(row), col+3)):
                    cell_val = str(row.iloc[c]) if pd.notna(row.iloc[c]) else ""
                    if len(cell_val) > 20 and not any(x in cell_val for x in ["ИНН", "КПП", "Р/С"]):
                        fio = cell_val.strip()
                
                # Ищем ИНН в этой же строке
                for c in range(len(row)):
                    cell_val = str(row.iloc[c]) if pd.notna(row.iloc[c]) else ""
                    if "инн" in cell_val.lower():
                        if ":" in cell_val:
                            parts = cell_val.split(":")
                            if len(parts) > 1:
                                inn_candidate = parts[1].strip()
                                if inn_candidate.isdigit() and len(inn_candidate) >= 10:
                                    inn = inn_candidate
                        elif c + 1 < len(row) and pd.notna(row.iloc[c + 1]):
                            inn_candidate = str(row.iloc[c + 1]).strip()
                            if inn_candidate.isdigit() and len(inn_candidate) >= 10:
                                inn = inn_candidate
                break
            
            # Для выписок ОЗОН Банк
            if "клиент:" in val_lower:
                fio = val.replace("Клиент:", "").replace("ИП", "").strip()
                # Ищем ИНН в следующих строках
                if idx + 1 < len(df):
                    next_row = df.iloc[idx + 1]
                    for c in range(len(next_row)):
                        cell_val = str(next_row.iloc[c]) if pd.notna(next_row.iloc[c]) else ""
                        if "инн:" in cell_val.lower():
                            inn_candidate = cell_val.replace("ИНН:", "").strip()
                            if inn_candidate.isdigit() and len(inn_candidate) >= 10:
                                inn = inn_candidate
                                break
                break
    
    # Очищаем ФИО
    fio = fio.replace("Р/С:", "").replace("БИК:", "").strip()
    fio = ' '.join(fio.split())
    
    # Если ФИО не найдено, пытаемся извлечь из названия ИП
    if not fio:
        for idx, row in df.iterrows():
            for col in range(len(row)):
                val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
                if "ИП" in val and len(val) > 10:
                    fio = val.replace("ИП", "").strip()
                    break
            if fio:
                break
    
    return inn, fio

def extract_ip_accounts(df, ip_inn):
    """Извлекает счета ИП из выписки (только счета, принадлежащие ИП)"""
    accounts = []
    seen_numbers = set()
    
    # Ищем строку с номером счета ИП по ИНН
    ip_account_number = ""
    bank = ""
    bik = ""
    
    for idx, row in df.iterrows():
        for col in range(len(row)):
            val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
            # Ищем по ИНН
            if ip_inn and ip_inn in val:
                # Проверяем соседние колонки на номер счета
                for c in range(max(0, col-5), min(len(row), col+5)):
                    cell_val = str(row.iloc[c]) if pd.notna(row.iloc[c]) else ""
                    if "40802" in cell_val and len(cell_val) >= 20:
                        ip_account_number = ''.join(ch for ch in cell_val if ch.isdigit())
                        # Ищем банк и БИК
                        for r in range(max(0, idx-3), min(len(df), idx+4)):
                            for bc in range(max(0, c-5), min(len(row), c+8)):
                                bank_val = str(df.iloc[r, bc]) if pd.notna(df.iloc[r, bc]) else ""
                                if "БИК" in bank_val:
                                    bik = ''.join(ch for ch in bank_val if ch.isdigit())
                                    if len(bik) == 9:
                                        bik = bik
                                if any(x in bank_val for x in ["Банк", "БАНК", "ООО", "АО", "ПАО"]):
                                    if len(bank_val) > 3 and len(bank_val) < 100 and "БИК" not in bank_val:
                                        bank = bank_val.strip()
                        break
                break
        if ip_account_number:
            break
    
    # Если не нашли по ИНН, ищем по строке "Счет:"
    if not ip_account_number:
        for idx, row in df.iterrows():
            for col in range(len(row)):
                val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
                if "счет:" in val.lower():
                    for c in range(max(0, col-2), min(len(row), col+3)):
                        cell_val = str(row.iloc[c]) if pd.notna(row.iloc[c]) else ""
                        if "40802" in cell_val and len(cell_val) >= 20:
                            ip_account_number = ''.join(ch for ch in cell_val if ch.isdigit())
                            # Ищем банк и БИК
                            for r in range(max(0, idx-3), min(len(df), idx+4)):
                                for bc in range(max(0, c-5), min(len(row), c+8)):
                                    bank_val = str(df.iloc[r, bc]) if pd.notna(df.iloc[r, bc]) else ""
                                    if "БИК" in bank_val:
                                        bik = ''.join(ch for ch in bank_val if ch.isdigit())
                                        if len(bik) == 9:
                                            bik = bik
                                    if any(x in bank_val for x in ["Банк", "БАНК", "ООО", "АО", "ПАО"]):
                                        if len(bank_val) > 3 and len(bank_val) < 100 and "БИК" not in bank_val:
                                            bank = bank_val.strip()
                            break
                if ip_account_number:
                    break
            if ip_account_number:
                break
    
    # Добавляем найденный счет
    if ip_account_number and ip_account_number not in seen_numbers:
        accounts.append({
            'number': ip_account_number,
            'bank': bank,
            'bik': bik
        })
        seen_numbers.add(ip_account_number)
    
    return accounts

def parse_bank_statement(file_path):
    """Парсинг выписки: извлекаем доходы, данные ИП и счета"""
    df = pd.read_excel(file_path, header=None)
    
    # Извлекаем данные ИП
    ip_inn, ip_fio = extract_ip_data(df)
    
    # Извлекаем счета ИП
    ip_accounts = extract_ip_accounts(df, ip_inn)
    
    # Находим строку с заголовками (где есть "кредит")
    header_row = None
    credit_col = None
    date_col = None
    purpose_col = None
    
    for idx, row in df.iterrows():
        for col in range(len(row)):
            val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
            val_lower = val.lower()
            
            if "кредит" in val_lower or "по кредиту" in val_lower:
                header_row = idx
                credit_col = col
            if "дата" in val_lower:
                date_col = col
            if "назначение" in val_lower or "содержание" in val_lower:
                purpose_col = col
        
        if header_row is not None:
            break
    
    if header_row is None:
        raise Exception("Не найдена колонка 'Кредит'")
    
    # Берем данные после заголовка
    df_data = df.iloc[header_row + 1:].reset_index(drop=True)
    
    # Если не нашли колонку с датой, ищем первую колонку с датами
    if date_col is None:
        for col in range(len(df_data.columns)):
            for row in range(min(5, len(df_data))):
                val = str(df_data.iloc[row, col]) if pd.notna(df_data.iloc[row, col]) else ""
                if len(val) >= 8 and '.' in val:
                    try:
                        datetime.strptime(val, "%d.%m.%Y")
                        date_col = col
                        break
                    except:
                        pass
            if date_col is not None:
                break
    
    # Если не нашли колонку с назначением, берем последнюю
    if purpose_col is None:
        purpose_col = len(df_data.columns) - 1
    
    operations = []
    
    for idx, row in df_data.iterrows():
        try:
            # Сумма по кредиту
            credit_val = row.iloc[credit_col] if credit_col < len(row) else None
            if pd.isna(credit_val):
                continue
            
            amount = safe_float(credit_val)
            if amount <= 0:
                continue
            
            # Дата
            if date_col is None or date_col >= len(row):
                continue
            date_val = row.iloc[date_col]
            if pd.isna(date_val):
                continue
            date = parse_date(date_val)
            if not date:
                continue
            
            # Назначение
            purpose = ""
            if purpose_col < len(row):
                purpose_val = row.iloc[purpose_col]
                if pd.notna(purpose_val):
                    purpose = str(purpose_val)
            
            # Пропускаем строки "Итого"
            if "итого" in purpose.lower():
                continue
            
            # Исключаем переводы себе
            if any(word in purpose.lower() for word in ["собственных средств", "перевод собственных", "вывод собственных"]):
                continue
            
            # Номер документа
            doc_num = ""
            for col in range(min(5, len(row))):
                doc_val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
                if doc_val and doc_val != "nan" and not doc_val.replace('.', '').isdigit():
                    doc_num = doc_val
                    break
            
            operations.append({
                'date': date,
                'amount': amount,
                'purpose': purpose[:200],
                'document': f"{date.strftime('%d.%m.%Y')} {doc_num}" if doc_num else f"{date.strftime('%d.%m.%Y')} оп.{idx+1}"
            })
        except Exception as e:
            continue
    
    return operations, ip_inn, ip_fio, ip_accounts
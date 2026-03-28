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
    """Извлекает ИНН и ФИО из выписки (только из строк с данными ИП)"""
    inn = ""
    fio = ""
    
    # Ищем строку с данными ИП
    for idx, row in df.iterrows():
        for col in range(len(row)):
            val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
            val_lower = val.lower()
            
            # Ищем ФИО
            if "индивидуальный предприниматель" in val_lower:
                fio = val.replace("Индивидуальный предприниматель", "").replace("ИП", "").strip()
                # Проверяем соседние колонки
                if col + 1 < len(row) and pd.notna(row.iloc[col + 1]):
                    fio_candidate = str(row.iloc[col + 1]).strip()
                    if len(fio_candidate) > 10 and not fio_candidate.startswith("Р/С"):
                        fio = fio_candidate
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
    
    # Очищаем ФИО от лишних символов
    fio = fio.replace("Р/С:", "").replace("БИК:", "").strip()
    # Если ФИО слишком длинное и похоже на реквизиты, сбрасываем
    if len(fio) > 50 or "Р/С" in fio or "БИК" in fio:
        fio = ""
    
    return inn, fio

def extract_ip_accounts(df):
    """Извлекает счета ИП из выписки"""
    accounts = []
    seen_numbers = set()  # для уникальности счетов
    
    for idx, row in df.iterrows():
        for col in range(len(row)):
            val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
            
            # Ищем номер счета (обычно начинается с 40802 или 40817)
            if ("40802" in val or "40817" in val) and len(val) >= 20:
                # Очищаем номер счета от лишних символов
                account_number = val.strip()
                # Убираем возможные пробелы и лишние символы
                account_number = ''.join(c for c in account_number if c.isdigit())
                
                if account_number and account_number not in seen_numbers:
                    # Ищем банк и БИК
                    bank = ""
                    bik = ""
                    
                    # Проверяем строку вокруг
                    for c in range(max(0, col-5), min(len(row), col+5)):
                        bank_val = str(row.iloc[c]) if pd.notna(row.iloc[c]) else ""
                        bank_val_lower = bank_val.lower()
                        
                        # Ищем БИК
                        if "бик" in bank_val_lower:
                            if ":" in bank_val:
                                bik = bank_val.split(":")[-1].strip()
                            else:
                                # БИК может быть в соседней колонке
                                if c + 1 < len(row) and pd.notna(row.iloc[c + 1]):
                                    bik_candidate = str(row.iloc[c + 1]).strip()
                                    if len(bik_candidate) == 9 and bik_candidate.isdigit():
                                        bik = bik_candidate
                                    else:
                                        bik = bank_val.strip()
                        
                        # Ищем наименование банка
                        if any(x in bank_val_lower for x in ["банк", "банк"]):
                            if len(bank_val) > 3 and len(bank_val) < 100:
                                # Очищаем от лишнего
                                bank_val_clean = bank_val.replace("Р/С:", "").replace("р/с:", "").strip()
                                if "БИК" not in bank_val_clean:
                                    bank = bank_val_clean
                    
                    # Если не нашли банк, проверяем предыдущие/следующие строки
                    if not bank:
                        for r in range(max(0, idx-2), min(len(df), idx+3)):
                            for c in range(len(df.iloc[r])):
                                bank_val = str(df.iloc[r, c]) if pd.notna(df.iloc[r, c]) else ""
                                if any(x in bank_val.lower() for x in ["банк", "банк"]):
                                    if len(bank_val) > 3 and len(bank_val) < 100:
                                        bank = bank_val.strip()
                                        break
                            if bank:
                                break
                    
                    seen_numbers.add(account_number)
                    accounts.append({
                        'number': account_number,
                        'bank': bank,
                        'bik': bik
                    })
    
    return accounts

def parse_bank_statement(file_path):
    """Парсинг выписки: извлекаем доходы, данные ИП и счета"""
    df = pd.read_excel(file_path, header=None)
    
    # Извлекаем данные ИП
    ip_inn, ip_fio = extract_ip_data(df)
    
    # Извлекаем счета ИП
    ip_accounts = extract_ip_accounts(df)
    
    # Находим строку с заголовками (где есть "кредит")
    header_row = None
    credit_col = None
    date_col = None
    purpose_col = None
    
    for idx, row in df.iterrows():
        for col in range(len(row)):
            val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
            val_lower = val.lower()
            
            # Ищем колонку с кредитом
            if "кредит" in val_lower or "по кредиту" in val_lower:
                header_row = idx
                credit_col = col
            
            # Ищем колонку с датой
            if "дата" in val_lower:
                date_col = col
            
            # Ищем колонку с назначением
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
            
            # Номер документа (ищем в первых колонках)
            doc_num = ""
            for col in range(min(5, len(row))):
                doc_val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
                if doc_val and doc_val != "nan" and not doc_val.replace('.', '').isdigit():
                    doc_num = doc_val
                    break
            
            # Добавляем информацию о счете в операцию (для возможного использования)
            # Но сам счет ИП уже извлечен в ip_accounts
            
            operations.append({
                'date': date,
                'amount': amount,
                'purpose': purpose[:200],
                'document': f"{date.strftime('%d.%m.%Y')} {doc_num}" if doc_num else f"{date.strftime('%d.%m.%Y')} оп.{idx+1}"
            })
        except Exception as e:
            continue
    
    return operations, ip_inn, ip_fio, ip_accounts
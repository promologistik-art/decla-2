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
            
            # ВБ Банк: "Индивидуальный предприниматель"
            if "индивидуальный предприниматель" in val_lower:
                fio = val.replace("Индивидуальный предприниматель", "").replace("ИП", "").strip()
                if col + 1 < len(row) and pd.notna(row.iloc[col + 1]):
                    fio_candidate = str(row.iloc[col + 1]).strip()
                    if len(fio_candidate) > 10:
                        fio = fio_candidate
                
                for r in range(idx, min(idx + 5, len(df))):
                    for c in range(len(df.iloc[r])):
                        cell = str(df.iloc[r, c]) if pd.notna(df.iloc[r, c]) else ""
                        if "инн" in cell.lower():
                            if ":" in cell:
                                parts = cell.split(":")
                                if len(parts) > 1:
                                    inn_candidate = ''.join(ch for ch in parts[1] if ch.isdigit())
                                    if len(inn_candidate) == 12:
                                        inn = inn_candidate
                            elif c + 1 < len(df.iloc[r]) and pd.notna(df.iloc[r, c + 1]):
                                inn_candidate = ''.join(ch for ch in str(df.iloc[r, c + 1]) if ch.isdigit())
                                if len(inn_candidate) == 12:
                                    inn = inn_candidate
                break
            
            # ОЗОН Банк: "Клиент:"
            if "клиент:" in val_lower:
                fio = val.replace("Клиент:", "").replace("ИП", "").strip()
                for r in range(idx, min(idx + 3, len(df))):
                    for c in range(len(df.iloc[r])):
                        cell = str(df.iloc[r, c]) if pd.notna(df.iloc[r, c]) else ""
                        if "инн:" in cell.lower():
                            inn_candidate = ''.join(ch for ch in cell.replace("ИНН:", "") if ch.isdigit())
                            if len(inn_candidate) == 12:
                                inn = inn_candidate
                break
    
    fio = fio.replace("Р/С:", "").replace("БИК:", "").strip()
    fio = ' '.join(fio.split())
    
    return inn, fio

def extract_ip_accounts(df):
    """Извлекает счета ИП из выписки (по строке "Счет:")"""
    accounts = []
    seen_numbers = set()
    
    for idx, row in df.iterrows():
        for col in range(len(row)):
            val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
            if "счет:" in val.lower():
                if ":" in val:
                    account_number = ''.join(ch for ch in val.split(":")[-1] if ch.isdigit())
                elif col + 1 < len(row) and pd.notna(row.iloc[col + 1]):
                    account_number = ''.join(ch for ch in str(row.iloc[col + 1]) if ch.isdigit())
                else:
                    continue
                
                bank = ""
                bik = ""
                
                for r in range(max(0, idx-2), min(len(df), idx+3)):
                    for c in range(max(0, col-5), min(len(row), col+8)):
                        cell = str(df.iloc[r, c]) if pd.notna(df.iloc[r, c]) else ""
                        if "бик" in cell.lower():
                            bik = ''.join(ch for ch in cell if ch.isdigit())
                            if len(bik) == 9:
                                bik = bik
                        if any(x in cell for x in ["Банк", "БАНК", "ООО", "АО", "ПАО"]):
                            if len(cell) > 3 and len(cell) < 100 and "БИК" not in cell:
                                bank = cell.strip()
                
                if account_number and account_number not in seen_numbers and len(account_number) >= 20:
                    accounts.append({
                        'number': account_number,
                        'bank': bank,
                        'bik': bik
                    })
                    seen_numbers.add(account_number)
                break
    
    return accounts

def parse_bank_statement(file_path):
    """Парсинг выписки: извлекаем доходы, данные ИП и счета"""
    df = pd.read_excel(file_path, header=None)
    
    ip_inn, ip_fio = extract_ip_data(df)
    ip_accounts = extract_ip_accounts(df)
    
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
    
    df_data = df.iloc[header_row + 1:].reset_index(drop=True)
    
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
    
    if purpose_col is None:
        purpose_col = len(df_data.columns) - 1
    
    operations = []
    
    for idx, row in df_data.iterrows():
        try:
            credit_val = row.iloc[credit_col] if credit_col < len(row) else None
            if pd.isna(credit_val):
                continue
            
            amount = safe_float(credit_val)
            if amount <= 0:
                continue
            
            if date_col is None or date_col >= len(row):
                continue
            date_val = row.iloc[date_col]
            if pd.isna(date_val):
                continue
            date = parse_date(date_val)
            if not date:
                continue
            
            purpose = ""
            if purpose_col < len(row):
                purpose_val = row.iloc[purpose_col]
                if pd.notna(purpose_val):
                    purpose = str(purpose_val)
            
            if "итого" in purpose.lower():
                continue
            
            if any(word in purpose.lower() for word in ["собственных средств", "перевод собственных", "вывод собственных"]):
                continue
            
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
import pandas as pd
from datetime import datetime

def safe_float(val):
    try:
        if pd.isna(val):
            return 0.0
        if isinstance(val, str):
            return float(val.replace(" ", "").replace(",", "."))
        return float(val)
    except:
        return 0.0

def parse_date(val):
    if isinstance(val, datetime):
        return val
    if isinstance(val, pd.Timestamp):
        return val.to_pydatetime()
    if isinstance(val, str):
        try:
            return datetime.strptime(val.split()[0], '%Y-%m-%d')
        except:
            pass
    return None

def detect_tax_object(df, col_op, col_kbk):
    """Определяет объект налогообложения по КБК авансовых платежей"""
    kbk_6 = '18210501011011000110'   # УСН Доходы 6%
    kbk_15 = '18210501021011000110'  # УСН Доходы минус расходы 15%
    
    has_6 = False
    has_15 = False
    
    for _, row in df.iterrows():
        op = str(row.get(col_op, '')).lower()
        kbk = str(row.get(col_kbk, ''))
        
        if 'уплата' in op or 'платеж' in op:
            if kbk_6 in kbk:
                has_6 = True
            if kbk_15 in kbk:
                has_15 = True
    
    if has_6 and not has_15:
        return 1  # Доходы
    elif has_15 and not has_6:
        return 2  # Доходы минус расходы
    else:
        return None  # Не определилось, нужно спросить

def parse_ens_statement(file_path):
    """Парсинг CSV выписки ЕНС"""
    df = None
    for sep in [';', ',']:
        try:
            df = pd.read_csv(file_path, sep=sep, encoding='utf-8')
            if len(df.columns) > 1:
                break
        except:
            try:
                df = pd.read_csv(file_path, sep=sep, encoding='windows-1251')
                if len(df.columns) > 1:
                    break
            except:
                continue
    
    if df is None or len(df.columns) <= 1:
        raise Exception("Не удалось прочитать файл")
    
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    result = {
        'insurance_accrued': 0.0,
        'insurance_paid': 0.0,
        'insurance_paid_dates': [],
        'penalties': 0.0,
        'oktmo': '',
        'usn_payments': [],
        'tax_object': None  # 1 - Доходы, 2 - Доходы минус расходы
    }
    
    col_op = None
    col_amount = None
    col_date = None
    col_kbk = None
    col_oktmo = None
    
    for col in df.columns:
        if 'операции' in col:
            col_op = col
        elif 'сумма' in col:
            col_amount = col
        elif 'дата' in col:
            col_date = col
        elif 'кбк' in col:
            col_kbk = col
        elif 'октмо' in col:
            col_oktmo = col
    
    if col_op is None:
        col_op = df.columns[0]
    if col_amount is None:
        for col in df.columns:
            if df[col].dtype in ['float64', 'int64']:
                col_amount = col
                break
    
    # Определяем объект налогообложения
    result['tax_object'] = detect_tax_object(df, col_op, col_kbk)
    
    # Ищем ОКТМО в операциях уплаты
    found_oktmo = None
    for _, row in df.iterrows():
        op = str(row.get(col_op, '')).lower()
        if col_oktmo:
            oktmo_val = row.get(col_oktmo)
            if pd.notna(oktmo_val) and str(oktmo_val).strip():
                if 'уплата' in op or 'платеж' in op:
                    found_oktmo = str(oktmo_val).strip()
                    break
    
    if not found_oktmo and col_oktmo:
        for _, row in df.iterrows():
            oktmo_val = row.get(col_oktmo)
            if pd.notna(oktmo_val) and str(oktmo_val).strip():
                found_oktmo = str(oktmo_val).strip()
                break
    
    result['oktmo'] = found_oktmo if found_oktmo else ""
    
    # Парсинг остальных данных
    for _, row in df.iterrows():
        try:
            op = str(row.get(col_op, '')).lower()
            amount = safe_float(row.get(col_amount, 0))
            kbk = str(row.get(col_kbk, ''))
            date = parse_date(row.get(col_date, ''))
            
            if 'начислено' in op and 'страховые взносы' in op:
                result['insurance_accrued'] += abs(amount)
            elif 'пеня' in op:
                result['penalties'] += abs(amount)
            elif 'уплата' in op or 'платеж' in op:
                # КБК УСН (авансовые платежи) - оба варианта
                if '1821050101' in kbk:  # УСН
                    if date and amount > 0:
                        result['usn_payments'].append({
                            'date': date,
                            'amount': amount,
                            'kbk': kbk
                        })
                # Страховые взносы
                elif date and date.year == 2026 and amount > 0:
                    result['insurance_paid'] += amount
                    result['insurance_paid_dates'].append(date)
            elif '18210202000010000160' in kbk and amount > 0:
                if date and date.year == 2026:
                    result['insurance_paid'] += amount
                    if date not in result['insurance_paid_dates']:
                        result['insurance_paid_dates'].append(date)
        except:
            continue
    
    return result
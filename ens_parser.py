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
        'usn_payments': []
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
    
    for _, row in df.iterrows():
        if col_oktmo:
            oktmo_val = row.get(col_oktmo)
            if pd.notna(oktmo_val) and str(oktmo_val).strip():
                result['oktmo'] = str(oktmo_val).strip()
                break
    
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
                if '18201061201010000510' in kbk:
                    if date and amount > 0:
                        result['usn_payments'].append({
                            'date': date,
                            'amount': amount
                        })
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
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
        for fmt in ["%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S"]:
            try:
                return datetime.strptime(val.strip(), fmt)
            except:
                continue
    return None

def is_income(text):
    text = str(text).lower()
    exclude = ["собственных средств", "перевод собственных", "вывод собственных", "комиссия"]
    for w in exclude:
        if w in text:
            return False
    keywords = ["оплата за товар", "оплата по договору", "оплата за услуги", 
                "интернет решения", "озон", "по реестру", "за товар"]
    for w in keywords:
        if w in text:
            return True
    return False

def parse_bank_statement(file_path):
    df = pd.read_excel(file_path, header=None)
    
    # Ищем строку с данными (первая строка, где первая ячейка похожа на дату)
    data_start = None
    for idx, row in df.iterrows():
        first = str(row.iloc[0]) if len(row) > 0 else ""
        if first and len(first) >= 8 and '.' in first and first.replace('.', '').isdigit():
            data_start = idx
            break
    
    if data_start is None:
        raise Exception("Не удалось найти строки с данными")
    
    df_data = df.iloc[data_start:].reset_index(drop=True)
    
    # Определяем колонки по первым 5 строкам
    col_date = None
    col_amount = None
    col_text = None
    
    for col in range(len(df_data.columns)):
        sample = []
        for row in range(min(5, len(df_data))):
            val = df_data.iloc[row, col]
            if pd.notna(val):
                sample.append(str(val))
        
        if not sample:
            continue
        
        sample_str = ' '.join(sample)
        
        # Дата
        if any('.' in s and len(s) >= 8 and s.replace('.', '').isdigit() for s in sample):
            if col_date is None:
                col_date = col
        
        # Сумма (число)
        elif any(s.replace('.', '').replace('-', '').isdigit() and len(s) > 0 for s in sample):
            if col_amount is None:
                col_amount = col
        
        # Текст
        elif any(len(s) > 20 for s in sample):
            if col_text is None:
                col_text = col
    
    if col_date is None:
        raise Exception("Не найдена колонка с датой")
    
    operations = []
    
    for idx, row in df_data.iterrows():
        try:
            date_val = row.iloc[col_date]
            if pd.isna(date_val):
                continue
            date = parse_date(date_val)
            if not date:
                continue
            
            amount = 0.0
            if col_amount is not None:
                amount = safe_float(row.iloc[col_amount])
            
            if amount <= 0:
                continue
            
            text = ""
            if col_text is not None:
                text = str(row.iloc[col_text])
            
            if is_income(text):
                operations.append({
                    'date': date,
                    'amount': amount,
                    'purpose': text[:200],
                    'document': f"{date.strftime('%d.%m.%Y')} п/п {idx+1}"
                })
        except:
            continue
    
    return operations
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
        formats = ["%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S"]
        for fmt in formats:
            try:
                return datetime.strptime(val, fmt)
            except:
                continue
    return None

def is_income(text):
    """Доход — всё, кроме переводов себе, комиссий и возвратов"""
    text = str(text).lower()
    
    # Не доход
    exclude = [
        "собственных средств", "перевод собственных", "вывод собственных",
        "комиссия", "возврат"
    ]
    for w in exclude:
        if w in text:
            return False
    
    # Всё остальное — доход
    return True

def parse_bank_statement(file_path):
    """Парсинг выписки: ищем колонку 'Кредит' или 'По кредиту'"""
    df = pd.read_excel(file_path, header=None)
    
    # 1. Находим строку с заголовками (где есть "кредит")
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
    
    # 2. Берем данные после заголовка
    df_data = df.iloc[header_row + 1:].reset_index(drop=True)
    
    # 3. Если не нашли колонку с датой, ищем первую колонку с датами
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
    
    # 4. Если не нашли колонку с назначением, берем последнюю
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
            date_val = row.iloc[date_col] if date_col < len(row) else None
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
            
            # Исключаем строки с "Итого"
            if "итого" in purpose.lower():
                continue
            
            # Если доход
            if is_income(purpose):
                operations.append({
                    'date': date,
                    'amount': amount,
                    'purpose': purpose[:200],
                    'document': f"{date.strftime('%d.%m.%Y')} оп.{idx+1}"
                })
                
        except Exception as e:
            continue
    
    return operations
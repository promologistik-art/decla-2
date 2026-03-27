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
        # Очищаем строку
        val = val.strip()
        formats = ["%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M"]
        for fmt in formats:
            try:
                return datetime.strptime(val, fmt)
            except:
                continue
    return None

def is_income(text):
    text = str(text).lower()
    exclude = ["собственных средств", "перевод собственных", "вывод собственных", "комиссия", "уплата налога"]
    for w in exclude:
        if w in text:
            return False
    keywords = ["оплата за товар", "оплата по договору", "оплата за услуги", 
                "интернет решения", "озон", "по реестру", "за товар", 
                "оплата за ооо", "платеж по ден.треб"]
    for w in keywords:
        if w in text:
            return True
    return False

def parse_bank_statement(file_path):
    """Универсальный парсинг выписки"""
    df = pd.read_excel(file_path, header=None)
    
    # Ищем строку с данными (где есть дата в формате дд.мм.гггг в любой колонке)
    data_start = None
    date_col = None
    
    for idx, row in df.iterrows():
        for col in range(len(row)):
            val = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
            # Проверяем, похоже ли значение на дату
            if len(val) >= 8 and '.' in val:
                parts = val.split('.')
                if len(parts) == 3 and len(parts[0]) <= 2 and len(parts[1]) <= 2 and len(parts[2]) == 4:
                    try:
                        datetime.strptime(val, "%d.%m.%Y")
                        data_start = idx
                        date_col = col
                        break
                    except:
                        pass
        if data_start is not None:
            break
    
    if data_start is None:
        raise Exception("Не удалось найти строки с датами")
    
    # Берем данные с найденной строки
    df_data = df.iloc[data_start:].reset_index(drop=True)
    
    # Определяем колонки по первым строкам
    col_date = date_col
    col_amount = None
    col_text = None
    
    # Ищем колонку с суммой (кредит/дебет)
    for col in range(len(df_data.columns)):
        sample = []
        for row in range(min(5, len(df_data))):
            val = df_data.iloc[row, col]
            if pd.notna(val):
                sample.append(str(val))
        
        if not sample:
            continue
        
        # Пробуем найти числовые значения (суммы)
        for s in sample:
            try:
                num = safe_float(s)
                if num != 0 and col != col_date:
                    col_amount = col
                    break
            except:
                pass
        if col_amount is not None:
            break
    
    # Ищем колонку с текстом (назначение платежа)
    for col in range(len(df_data.columns)):
        if col == col_date or col == col_amount:
            continue
        sample = []
        for row in range(min(5, len(df_data))):
            val = df_data.iloc[row, col]
            if pd.notna(val):
                sample.append(str(val))
        
        if sample and any(len(s) > 30 for s in sample):
            col_text = col
            break
    
    # Если не нашли текстовую колонку, берем первую не дату и не сумму
    if col_text is None:
        for col in range(len(df_data.columns)):
            if col != col_date and col != col_amount:
                col_text = col
                break
    
    operations = []
    
    for idx, row in df_data.iterrows():
        try:
            # Дата
            date_val = row.iloc[col_date]
            if pd.isna(date_val):
                continue
            date = parse_date(date_val)
            if not date:
                continue
            
            # Сумма
            amount = 0.0
            if col_amount is not None:
                amount = safe_float(row.iloc[col_amount])
            
            if amount <= 0:
                continue
            
            # Текст
            text = ""
            if col_text is not None:
                text = str(row.iloc[col_text])
            
            if is_income(text):
                operations.append({
                    'date': date,
                    'amount': amount,
                    'purpose': text[:200],
                    'document': f"{date.strftime('%d.%m.%Y')} оп.{idx+1}"
                })
        except Exception as e:
            continue
    
    return operations
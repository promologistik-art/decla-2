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
            
            # Ищем ИНН
            if "инн" in val_lower and not inn:
                # ИНН может быть в соседней колонке
                if col + 1 < len(row) and pd.notna(row.iloc[col + 1]):
                    inn_candidate = str(row.iloc[col + 1]).strip()
                    if inn_candidate.isdigit() and len(inn_candidate) >= 10:
                        inn = inn_candidate
                # Или в текущей колонке после ":"
                if ":" in val:
                    parts = val.split(":")
                    if len(parts) > 1:
                        inn_candidate = parts[1].strip()
                        if inn_candidate.isdigit() and len(inn_candidate) >= 10:
                            inn = inn_candidate
            
            # Ищем ФИО
            if "клиент:" in val_lower or "индивидуальный предприниматель" in val_lower:
                # ФИО в текущей колонке
                if "ип" in val_lower:
                    fio = val.replace("ИП", "").strip()
                else:
                    fio = val.split(":", 1)[-1].strip() if ":" in val else val
                # Проверяем, есть ли ФИО в соседней колонке
                if col + 1 < len(row) and pd.notna(row.iloc[col + 1]):
                    fio_candidate = str(row.iloc[col + 1]).strip()
                    if len(fio_candidate) > 10:
                        fio = fio_candidate
    
    # Очищаем ФИО
    fio = fio.replace("ИП", "").replace("индивидуальный предприниматель", "").strip()
    
    return inn, fio

def parse_bank_statement(file_path):
    """Парсинг выписки: извлекаем доходы и данные ИП"""
    df = pd.read_excel(file_path, header=None)
    
    # Извлекаем данные ИП
    ip_inn, ip_fio = extract_ip_data(df)
    
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
            
            operations.append({
                'date': date,
                'amount': amount,
                'purpose': purpose[:200],
                'document': f"{date.strftime('%d.%m.%Y')} {doc_num}" if doc_num else f"{date.strftime('%d.%m.%Y')} оп.{idx+1}"
            })
        except Exception as e:
            continue
    
    return operations, ip_inn, ip_fio
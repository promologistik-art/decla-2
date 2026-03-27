#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import tempfile
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
ADMIN_IDS = [int(id) for id in os.getenv("ADMIN_IDS", "").split(",") if id]

# Папки
DATA_DIR = "data"
OUTPUT_DIR = "output"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Ключевые слова для определения дохода
INCOME_KEYWORDS = [
    "оплата за товар", "оплата по договору", "оплата за услуги",
    "интернет решения", "озон", "по реестру", "оплата по контракту",
    "платеж по ден.треб", "за товар"
]

# Исключаемые операции
EXCLUDE_KEYWORDS = [
    "собственных средств", "перевод собственных", "вывод собственных",
    "комиссия", "уплата налога", "страховые взносы"
]

# Хранилище сессий пользователей
user_sessions = {}


class UserSession:
    def __init__(self, user_id):
        self.user_id = user_id
        self.bank_operations = []  # все доходы из банков
        self.ens_data = {
            'insurance_accrued': 0,
            'insurance_paid': 0,
            'insurance_paid_dates': [],
            'penalties': 0
        }

    def add_bank_operations(self, operations):
        self.bank_operations.extend(operations)

    def set_ens_data(self, data):
        self.ens_data = data

    def reset(self):
        self.bank_operations = []
        self.ens_data = {
            'insurance_accrued': 0,
            'insurance_paid': 0,
            'insurance_paid_dates': [],
            'penalties': 0
        }


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========

def safe_float(value):
    try:
        if pd.isna(value):
            return 0.0
        if isinstance(value, str):
            cleaned = value.replace(" ", "").replace(",", ".")
            return float(cleaned)
        return float(value)
    except:
        return 0.0


def parse_date(date_str):
    if isinstance(date_str, datetime):
        return date_str
    if isinstance(date_str, str):
        formats = ["%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S"]
        for fmt in formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except:
                continue
    return None


def is_income(purpose):
    purpose_lower = str(purpose).lower()
    for word in EXCLUDE_KEYWORDS:
        if word in purpose_lower:
            return False
    for word in INCOME_KEYWORDS:
        if word in purpose_lower:
            return True
    return False


# ========== ПАРСИНГ ВЫПИСКИ ИЗ БАНКА ==========

def parse_bank_statement(file_path):
    """Парсинг Excel-выписки, возвращает список доходов"""
    try:
        df = pd.read_excel(file_path, header=None)
        
        # Ищем строку с заголовками
        header_row = None
        for idx, row in df.iterrows():
            row_str = ' '.join(str(v) for v in row.values if pd.notna(v)).lower()
            if 'дата' in row_str and ('сумма' in row_str or 'кредит' in row_str):
                header_row = idx
                break
        
        if header_row is not None:
            headers = df.iloc[header_row].values
            df.columns = headers
            df = df.iloc[header_row + 1:].reset_index(drop=True)
        
        # Определяем колонки
        col_date = None
        col_credit = None
        col_debit = None
        col_purpose = None
        
        for col in df.columns:
            col_str = str(col).lower()
            if 'дата' in col_str:
                col_date = col
            elif 'кредит' in col_str or 'поступление' in col_str or 'приход' in col_str:
                col_credit = col
            elif 'дебет' in col_str or 'списание' in col_str or 'расход' in col_str:
                col_debit = col
            elif 'назначение' in col_str or 'содержание' in col_str or 'назначение платежа' in col_str:
                col_purpose = col
        
        if col_date is None:
            return []
        
        operations = []
        for idx, row in df.iterrows():
            date = parse_date(row.get(col_date))
            if not date:
                continue
            
            purpose = str(row.get(col_purpose, ''))
            
            # Ищем сумму
            amount = 0
            if col_credit:
                amount = safe_float(row.get(col_credit, 0))
            if amount == 0 and col_debit:
                amount = safe_float(row.get(col_debit, 0))
                if amount > 0:
                    amount = -amount
            
            if amount > 0 and is_income(purpose):
                operations.append({
                    'date': date,
                    'amount': amount,
                    'purpose': purpose[:200],
                    'document': f"п/п {idx+1}"
                })
        
        return operations
        
    except Exception as e:
        raise Exception(f"Ошибка парсинга: {e}")


# ========== ПАРСИНГ ВЫПИСКИ ЕНС ==========

def parse_ens_statement(file_path):
    """Парсинг CSV выписки ЕНС (формат ФНС России)"""
    try:
        df = None
        
        # Пробуем разные кодировки и разделители
        encodings = ['utf-8', 'windows-1251', 'cp1251']
        separators = [';', ',', '\t']
        
        for enc in encodings:
            for sep in separators:
                try:
                    df = pd.read_csv(file_path, sep=sep, encoding=enc, on_bad_lines='skip')
                    if len(df.columns) > 1:
                        break
                except:
                    continue
            if df is not None and len(df.columns) > 1:
                break
        
        if df is None or len(df.columns) <= 1:
            raise Exception("Не удалось определить формат файла")
        
        result = {
            'insurance_accrued': 0.0,
            'insurance_paid': 0.0,
            'insurance_paid_dates': [],
            'penalties': 0.0
        }
        
        # Приводим названия колонок к нижнему регистру
        df.columns = [str(col).strip().lower() for col in df.columns]
        
        # Ищем нужные колонки
        col_operation = None
        col_amount = None
        col_date = None
        col_obligation = None
        col_kbk = None
        
        for col in df.columns:
            if 'наименование операции' in col or 'операция' in col:
                col_operation = col
            elif 'сумма' in col and 'операции' in col:
                col_amount = col
            elif 'дата' in col:
                col_date = col
            elif 'наименование обязательства' in col or 'обязательство' in col:
                col_obligation = col
            elif 'кбк' in col:
                col_kbk = col
        
        # Если не нашли колонку суммы, ищем любую числовую колонку
        if col_amount is None:
            for col in df.columns:
                if df[col].dtype in ['float64', 'int64']:
                    col_amount = col
                    break
        
        # Если не нашли колонку операций, берем первую
        if col_operation is None and len(df.columns) > 0:
            col_operation = df.columns[0]
        
        print(f"[DEBUG] Колонки: operation={col_operation}, amount={col_amount}, date={col_date}")
        
        for idx, row in df.iterrows():
            try:
                # Получаем данные
                operation = str(row.get(col_operation, '')).lower() if col_operation else ''
                obligation = str(row.get(col_obligation, '')).lower() if col_obligation else ''
                kbk = str(row.get(col_kbk, '')) if col_kbk else ''
                
                # Получаем сумму
                amount = 0.0
                if col_amount:
                    val = row.get(col_amount)
                    if pd.notna(val):
                        amount = safe_float(val)
                
                # Получаем дату
                date_obj = None
                if col_date:
                    date_str = str(row.get(col_date, ''))
                    if date_str and date_str != 'nan':
                        date_obj = parse_date(date_str)
                
                # Начисление страховых взносов
                if ('начислено' in operation or 'начисление' in operation) and \
                   ('страховые взносы' in obligation or 'фиксированный размер' in operation):
                    result['insurance_accrued'] += abs(amount)
                
                # Пени
                elif 'пеня' in operation or 'пени' in operation:
                    result['penalties'] += abs(amount)
                
                # Уплата (ЕНП)
                elif 'уплата' in operation or 'платеж' in operation:
                    # Если дата в 2026 году — это уплата взносов за 2025
                    if date_obj and date_obj.year == 2026 and amount > 0:
                        result['insurance_paid'] += amount
                        result['insurance_paid_dates'].append(date_obj)
                
                # Проверка по КБК страховых взносов
                if '18210202000010000160' in kbk and 'уплата' in operation:
                    if date_obj and date_obj.year == 2026:
                        result['insurance_paid'] += amount
                        if date_obj not in result['insurance_paid_dates']:
                            result['insurance_paid_dates'].append(date_obj)
                        
            except Exception as e:
                print(f"[DEBUG] Ошибка строки {idx}: {e}")
                continue
        
        return result
        
    except Exception as e:
        raise Exception(f"Ошибка парсинга ЕНС: {e}")


# ========== ГЕНЕРАЦИЯ КУДиР ==========

def generate_kudir(operations, output_path):
    """Генерация КУДиР в Excel"""
    if not operations:
        return 0
    
    sorted_ops = sorted(operations, key=lambda x: x['date'])
    
    wb = Workbook()
    ws = wb.active
    ws.title = "КУДиР"
    
    # Заголовки
    ws['A1'] = "Книга учета доходов и расходов ИП на УСН"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:D1')
    
    ws['A2'] = "за 2025 год"
    ws.merge_cells('A2:D2')
    
    headers = ['№ п/п', 'Дата и номер документа', 'Содержание операции', 'Сумма дохода']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    # Данные
    total = 0
    for idx, op in enumerate(sorted_ops, 1):
        ws.cell(row=idx + 4, column=1, value=idx)
        ws.cell(row=idx + 4, column=2, value=f"{op['date'].strftime('%d.%m.%Y')} {op['document']}")
        ws.cell(row=idx + 4, column=3, value=op['purpose'])
        ws.cell(row=idx + 4, column=4, value=f"{op['amount']:,.2f}".replace(",", " "))
        total += op['amount']
    
    # Итог
    total_row = len(sorted_ops) + 5
    ws.cell(row=total_row, column=3, value="ИТОГО:")
    ws.cell(row=total_row, column=3).font = Font(bold=True)
    ws.cell(row=total_row, column=4, value=f"{total:,.2f}".replace(",", " "))
    ws.cell(row=total_row, column=4).font = Font(bold=True)
    
    # Ширина колонок
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 15
    
    wb.save(output_path)
    return total


# ========== ГЕНЕРАЦИЯ ДЕКЛАРАЦИИ ==========

def generate_declaration(operations, ens_data, output_excel, output_xml):
    """Генерация декларации (Excel и XML)"""
    
    # Расчет доходов по кварталам
    quarterly = {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0}
    for op in operations:
        quarter = (op['date'].month - 1) // 3 + 1
        quarterly[quarter] += op['amount']
    
    total_income = sum(quarterly.values())
    tax_rate = 6
    tax_amount = total_income * tax_rate / 100
    
    # Проверяем, были ли уплачены взносы в 2025 году
    paid_in_2025 = any(d.year == 2025 for d in ens_data.get('insurance_paid_dates', []))
    if paid_in_2025:
        tax_payable = max(0, tax_amount - ens_data.get('insurance_paid', 0))
    else:
        tax_payable = tax_amount
    
    # Квартальные суммы нарастающим
    cum_income = {
        1: quarterly[1],
        2: quarterly[1] + quarterly[2],
        3: quarterly[1] + quarterly[2] + quarterly[3],
        4: total_income
    }
    
    # ========== Excel ==========
    wb = Workbook()
    ws = wb.active
    ws.title = "Декларация УСН"
    
    ws['A1'] = "Налоговая декларация по УСН за 2025 год"
    ws['A1'].font = Font(bold=True, size=12)
    ws.merge_cells('A1:C1')
    
    # Раздел 2.1.1
    ws['A3'] = "Раздел 2.1.1. Доходы"
    ws['A3'].font = Font(bold=True)
    
    data = [
        ("Доход за 1 квартал", "110", cum_income[1]),
        ("Доход за полугодие", "111", cum_income[2]),
        ("Доход за 9 месяцев", "112", cum_income[3]),
        ("Доход за год", "113", cum_income[4]),
        ("Налоговая ставка (%)", "120", tax_rate),
        ("Сумма налога за 1 квартал", "130", cum_income[1] * tax_rate / 100),
        ("Сумма налога за полугодие", "131", cum_income[2] * tax_rate / 100),
        ("Сумма налога за 9 месяцев", "132", cum_income[3] * tax_rate / 100),
        ("Сумма налога за год", "133", tax_amount),
    ]
    
    for idx, (name, code, val) in enumerate(data, 5):
        ws.cell(row=idx, column=1, value=name)
        ws.cell(row=idx, column=2, value=code)
        ws.cell(row=idx, column=3, value=round(val, 2))
    
    # Раздел 1.1
    start = 5 + len(data) + 2
    ws.cell(row=start, column=1, value="Раздел 1.1. Сумма налога к уплате")
    ws.cell(row=start, column=1).font = Font(bold=True)
    
    tax_data = [
        ("Код ОКТМО", "010", "36701320"),
        ("Аванс к уплате за 1 квартал", "020", 0),
        ("Аванс к уплате за полугодие", "040", 0),
        ("Аванс к уплате за 9 месяцев", "070", 0),
        ("Налог к уплате за год", "100", round(tax_payable, 2)),
    ]
    
    for idx, (name, code, val) in enumerate(tax_data, start + 2):
        ws.cell(row=idx, column=1, value=name)
        ws.cell(row=idx, column=2, value=code)
        ws.cell(row=idx, column=3, value=val)
    
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 20
    
    wb.save(output_excel)
    
    # ========== XML ==========
    xml_content = f'''<?xml version="1.0" encoding="UTF-8"?>
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
        <ИНН>632312967829</ИНН>
        <ИП>
            <ФИО>
                <Фамилия>Леонтьев</Фамилия>
                <Имя>Артём</Имя>
                <Отчество>Владиславович</Отчество>
            </ФИО>
        </ИП>
    </Налогоплательщик>
    <Показатели>
        <Раздел1_1>
            <ОКТМО>36701320</ОКТМО>
            <СумАван010>0</СумАван010>
            <СумАван020>0</СумАван020>
            <СумАван040>0</СумАван040>
            <СумАван070>0</СумАван070>
            <СумНал100>{int(tax_payable)}</СумНал100>
        </Раздел1_1>
        <Раздел2_1_1>
            <СумДоход110>{int(cum_income[1])}</СумДоход110>
            <СумДоход111>{int(cum_income[2])}</СумДоход111>
            <СумДоход112>{int(cum_income[3])}</СумДоход112>
            <СумДоход113>{int(cum_income[4])}</СумДоход113>
            <НалСтавка120>{tax_rate}</НалСтавка120>
            <СумИсчисНал130>{int(cum_income[1] * tax_rate / 100)}</СумИсчисНал130>
            <СумИсчисНал131>{int(cum_income[2] * tax_rate / 100)}</СумИсчисНал131>
            <СумИсчисНал132>{int(cum_income[3] * tax_rate / 100)}</СумИсчисНал132>
            <СумИсчисНал133>{int(tax_amount)}</СумИсчисНал133>
            <СумУплНал140>0</СумУплНал140>
            <СумУплНал141>0</СумУплНал141>
            <СумУплНал142>0</СумУплНал142>
            <СумУплНал143>0</СумУплНал143>
        </Раздел2_1_1>
    </Показатели>
</Файл>'''
    
    with open(output_xml, 'w', encoding='utf-8') as f:
        f.write(xml_content)
    
    return tax_payable, total_income


# ========== ОБРАБОТЧИКИ TELEGRAM ==========

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_sessions[user_id] = UserSession(user_id)
    
    await update.message.reply_text(
        "🤖 *Бот для подготовки отчетности ИП на УСН*\n\n"
        "Я помогу вам:\n"
        "📊 Сформировать КУДиР на основе выписок из банков\n"
        "📝 Заполнить декларацию по УСН\n"
        "💰 Рассчитать налог к уплате\n\n"
        "*Как работать:*\n"
        "1️⃣ Загрузите выписки с расчетных счетов (Excel)\n"
        "2️⃣ Загрузите выписку с ЕНС (CSV)\n"
        "3️⃣ Введите /report\n\n"
        "📌 *Сроки за 2025 год:*\n"
        "• Декларацию сдать до *27 апреля 2026*\n"
        "• Налог уплатить до *28 апреля 2026*\n\n"
        "⚠️ Если не сдать декларацию в срок — налоговая может заблокировать счет.\n"
        "Просрочка уплаты налога счет не блокирует — только пени.",
        parse_mode="Markdown"
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in user_sessions:
        user_sessions[user_id] = UserSession(user_id)
    
    session = user_sessions[user_id]
    document = update.message.document
    filename = document.file_name.lower()
    
    file = await context.bot.get_file(document.file_id)
    
    with tempfile.NamedTemporaryFile(suffix=os.path.splitext(filename)[1], delete=False) as tmp:
        await file.download_to_drive(tmp.name)
        tmp_path = tmp.name
    
    try:
        if filename.endswith(('.xlsx', '.xls')):
            await update.message.reply_text("📥 Обрабатываю выписку из банка...")
            operations = parse_bank_statement(tmp_path)
            
            if operations:
                session.add_bank_operations(operations)
                total = sum(op['amount'] for op in operations)
                await update.message.reply_text(
                    f"✅ Найдено {len(operations)} операций\n"
                    f"💰 Сумма доходов: {total:,.2f} ₽\n\n"
                    f"Загружайте другие выписки или пришлите выписку с ЕНС (CSV)."
                )
            else:
                await update.message.reply_text(
                    "⚠️ В выписке не найдено доходов.\n"
                    "Проверьте, что в файле есть колонки с датой, суммой и назначением платежа."
                )
        
        elif filename.endswith('.csv'):
            await update.message.reply_text("📥 Обрабатываю выписку с ЕНС...")
            ens_data = parse_ens_statement(tmp_path)
            session.set_ens_data(ens_data)
            
            paid_in_2025 = any(d.year == 2025 for d in ens_data['insurance_paid_dates'])
            
            await update.message.reply_text(
                f"✅ Выписка ЕНС обработана!\n\n"
                f"📌 Страховые взносы:\n"
                f"• Начислено: {ens_data['insurance_accrued']:,.2f} ₽\n"
                f"• Уплачено: {ens_data['insurance_paid']:,.2f} ₽\n"
                f"• Уплачено в 2025: {'Да' if paid_in_2025 else 'Нет'}\n"
                f"• Пени: {ens_data['penalties']:,.2f} ₽\n\n"
                f"Теперь введите /report"
            )
        
        else:
            await update.message.reply_text("❌ Поддерживаются только .xlsx, .xls и .csv")
    
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")
    
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in user_sessions:
        await update.message.reply_text("Сначала загрузите выписки (/start)")
        return
    
    session = user_sessions[user_id]
    
    if not session.bank_operations:
        await update.message.reply_text("⚠️ Сначала загрузите выписки с расчетных счетов")
        return
    
    if not session.ens_data.get('insurance_accrued') and not session.ens_data.get('insurance_paid'):
        await update.message.reply_text("⚠️ Сначала загрузите выписку с ЕНС")
        return
    
    await update.message.reply_text("🔄 Формирую отчетность... Это может занять несколько секунд.")
    
    try:
        # Генерируем КУДиР
        kudir_path = os.path.join(OUTPUT_DIR, f"kudir_{user_id}.xlsx")
        total_income = generate_kudir(session.bank_operations, kudir_path)
        
        # Генерируем декларацию
        decl_excel = os.path.join(OUTPUT_DIR, f"declaration_{user_id}.xlsx")
        decl_xml = os.path.join(OUTPUT_DIR, f"declaration_{user_id}.xml")
        tax_payable, total_income = generate_declaration(
            session.bank_operations, session.ens_data, decl_excel, decl_xml
        )
        
        # Отправляем результат
        await update.message.reply_text(
            f"✅ *Отчетность готова!*\n\n"
            f"📊 *Доход за 2025:* {total_income:,.2f} ₽\n"
            f"💰 *Налог к уплате:* {tax_payable:,.2f} ₽\n\n"
            f"📌 *Сроки:*\n"
            f"• Декларацию сдать до *27 апреля 2026*\n"
            f"• Налог уплатить до *28 апреля 2026*\n\n"
            f"⚠️ Если не сдать декларацию в срок — заблокируют счет.\n"
            f"Просрочка уплаты налога счет не блокирует.\n\n"
            f"📎 Отправляю файлы...",
            parse_mode="Markdown"
        )
        
        with open(kudir_path, 'rb') as f:
            await update.message.reply_document(f, filename="КУДиР_2025.xlsx", caption="📘 Книга учета доходов и расходов")
        
        with open(decl_excel, 'rb') as f:
            await update.message.reply_document(f, filename="Декларация_УСН_2025.xlsx", caption="📝 Декларация (Excel для проверки)")
        
        with open(decl_xml, 'rb') as f:
            await update.message.reply_document(f, filename="declaration_usn_2025.xml", caption="📎 XML для загрузки в ЛК ФНС")
        
        await update.message.reply_text(
            "🎉 *Готово!*\n\n"
            "Что дальше:\n"
            "1. Проверьте декларацию в Excel\n"
            "2. Загрузите XML в Личный кабинет ИП на сайте ФНС\n"
            "3. Подпишите электронной подписью и отправьте\n"
            "4. Уплатите налог до 28 апреля 2026\n\n"
            "💡 *Важно:* взносы, уплаченные в 2026 году, не уменьшают налог за 2025. "
            "Они будут учтены при расчете налога за 2026 год."
        )
    
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")
        import traceback
        traceback.print_exc()


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id in user_sessions:
        user_sessions[user_id].reset()
        await update.message.reply_text("🔄 Данные сброшены. Начните с /start")
    else:
        await update.message.reply_text("Нет активной сессии. Используйте /start")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 *Помощь*\n\n"
        "*Команды:*\n"
        "/start — начать работу\n"
        "/report — сформировать отчетность\n"
        "/reset — сбросить все данные\n"
        "/help — эта справка\n\n"
        "*Файлы:*\n"
        "• Сначала загрузите Excel-выписки из банков\n"
        "• Затем загрузите CSV-выписку с ЕНС\n"
        "• Введите /report\n\n"
        "*Сроки за 2025 год:*\n"
        "• Декларация: до 27 апреля 2026\n"
        "• Уплата налога: до 28 апреля 2026\n\n"
        "*Важно:*\n"
        "• За несдачу декларации — блокировка счета\n"
        "• За просрочку уплаты налога — только пени",
        parse_mode="Markdown"
    )


def main():
    if not BOT_TOKEN:
        print("❌ Ошибка: BOT_TOKEN не задан в .env файле")
        sys.exit(1)
    
    app = Application.builder().token(BOT_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("🤖 Бот запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()
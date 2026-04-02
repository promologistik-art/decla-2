#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import tempfile
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

from bank_parser import parse_bank_statement
from ens_parser import parse_ens_statement
from report_generator import generate_report

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "")

DATA_DIR = "data"
OUTPUT_DIR = "output"
TEMPLATES_DIR = "templates"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

user_sessions = {}


class UserSession:
    def __init__(self, user_id):
        self.user_id = user_id
        self.bank_operations = []
        self.ens_data = {
            'insurance_accrued': 0,
            'insurance_paid': 0,
            'insurance_paid_dates': [],
            'penalties': 0,
            'usn_payments': []
        }
        self.ens_loaded = False
        self.inn = ""
        self.fio = ""
        self.oktmo = ""
        self.ip_accounts = []
        self.okved = ""
        self.phone = ""
        self.awaiting_okved = False
        self.awaiting_phone = False

    def add_bank_operations(self, operations, inn="", fio="", accounts=None):
        self.bank_operations.extend(operations)
        if inn and len(inn) >= 10 and inn.isdigit():
            self.inn = inn
        if fio and len(fio) > 10:
            self.fio = fio
        if accounts:
            for acc in accounts:
                if acc['number'] not in [a['number'] for a in self.ip_accounts]:
                    self.ip_accounts.append(acc)

    def set_ens_data(self, data):
        self.ens_data = data
        self.ens_loaded = True
        if 'oktmo' in data and data['oktmo']:
            self.oktmo = data['oktmo']

    def reset(self):
        self.bank_operations = []
        self.ens_data = {
            'insurance_accrued': 0,
            'insurance_paid': 0,
            'insurance_paid_dates': [],
            'penalties': 0,
            'usn_payments': []
        }
        self.ens_loaded = False
        self.inn = ""
        self.fio = ""
        self.oktmo = ""
        self.ip_accounts = []
        self.okved = ""
        self.phone = ""
        self.awaiting_okved = False
        self.awaiting_phone = False


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_sessions[user_id] = UserSession(user_id)
    
    await update.message.reply_text(
        "🤖 *Бот для подготовки отчетности ИП на УСН*\n\n"
        "1️⃣ Загрузите выписки с расчетных счетов (Excel)\n"
        "2️⃣ Загрузите выписку с ЕНС (CSV)\n"
        "3️⃣ Введите /report\n\n"
        "📌 *Сроки за 2025 год:*\n"
        "• Декларацию сдать до *27 апреля 2026*\n"
        "• Налог уплатить до *28 апреля 2026*",
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
            operations, inn, fio, accounts = parse_bank_statement(tmp_path)
            
            if operations:
                session.add_bank_operations(operations, inn, fio, accounts)
                total = sum(op['amount'] for op in operations)
                total_all = sum(op['amount'] for op in session.bank_operations)
                
                msg = f"✅ Найдено {len(operations)} операций\n💰 Сумма в файле: {total:,.2f} ₽\n📊 Всего загружено: {len(session.bank_operations)} операций на {total_all:,.2f} ₽"
                
                if session.inn:
                    msg += f"\n🏢 ИНН: {session.inn}"
                if session.fio:
                    msg += f"\n👤 ИП: {session.fio}"
                if session.ip_accounts:
                    msg += f"\n🏦 Счета: {', '.join([a['number'] for a in session.ip_accounts])}"
                
                await update.message.reply_text(msg)
                await update.message.reply_text(
                    "📌 *Следующий шаг:* загрузите выписку с Единого налогового счета (ЕНС) в формате CSV",
                    parse_mode="Markdown"
                )
            else:
                await update.message.reply_text("⚠️ В выписке не найдено доходов")
        
        elif filename.endswith('.csv'):
            await update.message.reply_text("📥 Обрабатываю выписку ЕНС...")
            ens_data = parse_ens_statement(tmp_path)
            session.set_ens_data(ens_data)
            
            paid_in_2025 = any(d.year == 2025 for d in ens_data['insurance_paid_dates'])
            oktmo = ens_data.get('oktmo', '')
            usn_payments = ens_data.get('usn_payments', [])
            
            await update.message.reply_text(
                f"✅ Выписка ЕНС обработана!\n\n"
                f"📌 Страховые взносы:\n"
                f"• Начислено: {ens_data['insurance_accrued']:,.2f} ₽\n"
                f"• Уплачено: {ens_data['insurance_paid']:,.2f} ₽\n"
                f"• Уплачено в 2025: {'Да' if paid_in_2025 else 'Нет'}\n"
                f"• Пени: {ens_data['penalties']:,.2f} ₽\n"
                f"• ОКТМО: {oktmo}\n"
                f"• Авансов по УСН: {len(usn_payments)}\n\n"
                f"✅ Теперь введите /report"
            )
        
        else:
            await update.message.reply_text("❌ Поддерживаются .xlsx, .xls, .csv")
    
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
        await update.message.reply_text("⚠️ Сначала загрузите выписки из банков")
        return
    
    if not session.ens_loaded:
        await update.message.reply_text("⚠️ Сначала загрузите выписку ЕНС")
        return
    
    if not session.okved:
        session.awaiting_okved = True
        await update.message.reply_text(
            "📝 Для заполнения декларации укажите код ОКВЭД\n"
            "Например: *4791* (торговля по почте/Интернет)\n\n"
            "Введите только цифры:",
            parse_mode="Markdown"
        )
        return
    
    if not session.phone:
        session.awaiting_phone = True
        await update.message.reply_text(
            "📞 Укажите контактный телефон\n"
            "Например: *89261234567*\n\n"
            "Введите номер:",
            parse_mode="Markdown"
        )
        return
    
    await update.message.reply_text("🔄 Формирую отчетность...")
    
    try:
        all_ops = []
        for op in session.bank_operations:
            if isinstance(op, dict):
                all_ops.append(op)
            elif isinstance(op, list):
                all_ops.extend(op)
        all_ops.sort(key=lambda x: x['date'])
        
        kudir_template = os.path.join(TEMPLATES_DIR, "KUDIR_template.xlsx")
        decl_template = os.path.join(TEMPLATES_DIR, "Declaration_template.xlsx")
        
        if not os.path.exists(kudir_template):
            await update.message.reply_text(f"❌ Шаблон КУДиР не найден")
            return
        
        if not os.path.exists(decl_template):
            await update.message.reply_text(f"❌ Шаблон декларации не найден")
            return
        
        inn = session.inn if session.inn else "632312967829"
        fio = session.fio if session.fio else "Леонтьев Артём Владиславович"
        oktmo = session.oktmo if session.oktmo else "36701320"
        ip_accounts = session.ip_accounts if session.ip_accounts else []
        okved = session.okved
        phone = session.phone
        
        if not ip_accounts:
            ip_accounts = [
                {'number': '40802810000000009773', 'bank': 'ООО "ВБ Банк"', 'bik': '044525450'},
                {'number': '40802810100000851604', 'bank': 'ООО "ОЗОН БАНК"', 'bik': '044525068'},
            ]
        
        kudir_path, decl_excel, decl_xml, total_income, tax_payable = generate_report(
            all_ops, session.ens_data, OUTPUT_DIR, user_id,
            kudir_template, decl_template, inn, fio, oktmo, ip_accounts, okved, phone
        )
        
        await update.message.reply_text(
            f"✅ *Отчетность готова!*\n\n"
            f"📊 Доход за 2025: {total_income:,.2f} ₽\n"
            f"💰 Налог к уплате: {tax_payable:,.2f} ₽\n\n"
            f"📌 Сдать декларацию до *27 апреля 2026*\n"
            f"📌 Уплатить налог до *28 апреля 2026*",
            parse_mode="Markdown"
        )
        
        with open(kudir_path, 'rb') as f:
            await update.message.reply_document(f, filename="КУДиР_2025.xlsx", caption="📘 Книга учета доходов и расходов")
        
        with open(decl_excel, 'rb') as f:
            await update.message.reply_document(f, filename="Декларация_УСН_2025.xlsx", caption="📝 Декларация по УСН")
        
        with open(decl_xml, 'rb') as f:
            await update.message.reply_document(f, filename="declaration_usn_2025.xml", caption="📎 XML для загрузки в ЛК ФНС")
    
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")
        import traceback
        traceback.print_exc()


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in user_sessions:
        await update.message.reply_text("Сначала загрузите выписки (/start)")
        return
    
    session = user_sessions[user_id]
    text = update.message.text.strip()
    
    if session.awaiting_okved:
        okved_digits = ''.join(ch for ch in text if ch.isdigit())
        if okved_digits:
            session.okved = okved_digits
            session.awaiting_okved = False
            await update.message.reply_text(f"✅ ОКВЭД сохранен: {okved_digits}\n\n📞 Теперь укажите контактный телефон:")
            session.awaiting_phone = True
        else:
            await update.message.reply_text("❌ Введите цифры ОКВЭД (например, 4791)")
        return
    
    if session.awaiting_phone:
        phone_digits = ''.join(ch for ch in text if ch.isdigit())
        if phone_digits:
            session.phone = phone_digits
            session.awaiting_phone = False
            await update.message.reply_text(f"✅ Телефон сохранен: {phone_digits}\n\n🔄 Введите /report")
        else:
            await update.message.reply_text("❌ Введите номер телефона цифрами")
        return


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id in user_sessions:
        user_sessions[user_id].reset()
        await update.message.reply_text("🔄 Данные сброшены")
    else:
        await update.message.reply_text("Нет активной сессии")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 *Помощь*\n\n"
        "/start — начать\n"
        "/report — сформировать отчетность\n"
        "/reset — сбросить данные\n"
        "/help — справка",
        parse_mode="Markdown"
    )


def main():
    if not BOT_TOKEN:
        print("❌ BOT_TOKEN не задан в .env")
        sys.exit(1)
    
    app = Application.builder().token(BOT_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    print("🤖 Бот запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()
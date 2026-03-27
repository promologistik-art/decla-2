#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import tempfile
from datetime import datetime
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
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

user_sessions = {}


class UserSession:
    def __init__(self, user_id):
        self.user_id = user_id
        self.bank_operations = []
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
            operations = parse_bank_statement(tmp_path)
            
            if operations:
                session.add_bank_operations(operations)
                total = sum(op['amount'] for op in operations)
                await update.message.reply_text(
                    f"✅ Найдено {len(operations)} операций\n💰 Сумма: {total:,.2f} ₽"
                )
            else:
                await update.message.reply_text("⚠️ Доходов не найдено")
        
        elif filename.endswith('.csv'):
            await update.message.reply_text("📥 Обрабатываю выписку ЕНС...")
            ens_data = parse_ens_statement(tmp_path)
            session.set_ens_data(ens_data)
            
            paid = any(d.year == 2025 for d in ens_data['insurance_paid_dates'])
            await update.message.reply_text(
                f"✅ Выписка ЕНС обработана\n"
                f"• Взносы уплачены в 2025: {'Да' if paid else 'Нет'}\n"
                f"• Сумма взносов: {ens_data['insurance_paid']:,.2f} ₽"
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
        await update.message.reply_text("⚠️ Загрузите выписки из банков")
        return
    
    if not session.ens_data.get('insurance_accrued') and not session.ens_data.get('insurance_paid'):
        await update.message.reply_text("⚠️ Загрузите выписку ЕНС")
        return
    
    await update.message.reply_text("🔄 Формирую отчетность...")
    
    try:
        kudir_path, decl_excel, decl_xml, total_income, tax_payable = generate_report(
            session.bank_operations, session.ens_data, OUTPUT_DIR, user_id
        )
        
        await update.message.reply_text(
            f"✅ *Отчетность готова*\n\n"
            f"📊 Доход: {total_income:,.2f} ₽\n"
            f"💰 Налог к уплате: {tax_payable:,.2f} ₽\n\n"
            f"📌 Сдать декларацию до *27 апреля 2026*",
            parse_mode="Markdown"
        )
        
        with open(kudir_path, 'rb') as f:
            await update.message.reply_document(f, filename="КУДиР_2025.xlsx")
        
        with open(decl_excel, 'rb') as f:
            await update.message.reply_document(f, filename="Декларация_УСН_2025.xlsx")
        
        with open(decl_xml, 'rb') as f:
            await update.message.reply_document(f, filename="declaration_usn_2025.xml")
    
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id in user_sessions:
        user_sessions[user_id].reset()
        await update.message.reply_text("🔄 Данные сброшены")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📌 *Команды:*\n"
        "/start — начать\n"
        "/report — сформировать отчетность\n"
        "/reset — сбросить данные\n"
        "/help — справка",
        parse_mode="Markdown"
    )


def main():
    if not BOT_TOKEN:
        print("❌ BOT_TOKEN не задан")
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
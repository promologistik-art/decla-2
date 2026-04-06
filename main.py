#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import json
import tempfile
from datetime import datetime, timedelta
from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackQueryHandler

from bank_parser import parse_bank_statement
from ens_parser import parse_ens_statement
from report_generator import generate_report

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
ADMIN_IDS = [int(x.strip()) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()]
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "silverzen")
SUBSCRIPTION_PRICE = int(os.getenv("SUBSCRIPTION_PRICE", "499"))

DATA_DIR = "data"
OUTPUT_DIR = "output"
TEMPLATES_DIR = "templates"
USERS_FILE = "users.json"

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

user_sessions = {}


def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def get_user_data(user_id):
    users = load_users()
    user_id_str = str(user_id)
    if user_id_str not in users:
        users[user_id_str] = {
            "demo_attempts": 0,
            "subscription_until": None,
            "created_at": datetime.now().isoformat(),
            "username": None,
            "first_name": None,
            "last_name": None
        }
        save_users(users)
    return users[user_id_str]


def update_user_data(user_id, **kwargs):
    users = load_users()
    user_id_str = str(user_id)
    if user_id_str not in users:
        users[user_id_str] = {}
    for key, value in kwargs.items():
        users[user_id_str][key] = value
    save_users(users)


def can_use_full_version(user_id):
    if is_admin(user_id):
        return True
    user_data = get_user_data(user_id)
    if user_data.get("subscription_until"):
        try:
            until = datetime.fromisoformat(user_data["subscription_until"])
            if datetime.now() < until:
                return True
        except:
            pass
    return False


def get_demo_attempts_left(user_id):
    user_data = get_user_data(user_id)
    attempts = user_data.get("demo_attempts", 0)
    return max(0, 3 - attempts)


def use_demo_attempt(user_id):
    user_data = get_user_data(user_id)
    attempts = user_data.get("demo_attempts", 0) + 1
    update_user_data(user_id, demo_attempts=attempts)
    return 3 - attempts


def is_admin(user_id):
    return user_id in ADMIN_IDS


def is_valid_fio(fio):
    if not fio:
        return False
    has_cyrillic = any('\u0400' <= c <= '\u04FF' for c in fio)
    is_only_digits = all(c.isdigit() or c.isspace() for c in fio)
    has_space = ' ' in fio
    return has_cyrillic and not is_only_digits and has_space


def detect_bank_name(filename):
    name_lower = filename.lower()
    if 'ozon' in name_lower:
        return 'ОЗОН Банк'
    elif 'vb' in name_lower or 'вб' in name_lower:
        return 'ВБ Банк'
    elif 'tinkoff' in name_lower or 'тинькофф' in name_lower:
        return 'Тинькофф'
    elif 'sber' in name_lower or 'сбер' in name_lower:
        return 'Сбербанк'
    elif 'alfa' in name_lower or 'альфа' in name_lower:
        return 'Альфа-Банк'
    else:
        return 'Банк'


class UserSession:
    def __init__(self, user_id):
        self.user_id = user_id
        self.bank_operations = []
        self.bank_files = []
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
        self.phone = ""
        self.awaiting_phone = False
        self.awaiting_oktmo = False
        self.awaiting_fio = False
        self.awaiting_inn = False

    def add_bank_operations(self, operations, bank_name="", inn="", fio="", accounts=None):
        self.bank_operations.extend(operations)
        self.bank_files.append(bank_name)
        
        if inn and len(inn) >= 10 and inn.isdigit() and not self.inn:
            self.inn = inn
        
        if is_valid_fio(fio) and not self.fio:
            self.fio = fio
        
        if accounts:
            for acc in accounts:
                if acc['number'] not in [a['number'] for a in self.ip_accounts]:
                    self.ip_accounts.append(acc)

    def set_ens_data(self, data):
        self.ens_data = data
        self.ens_loaded = True
        if 'oktmo' in data and data['oktmo']:
            oktmo_val = data['oktmo']
            if oktmo_val == "36701320":
                oktmo_val = "36701000"
            self.oktmo = oktmo_val

    def reset(self):
        self.bank_operations = []
        self.bank_files = []
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
        self.phone = ""
        self.awaiting_phone = False
        self.awaiting_oktmo = False
        self.awaiting_fio = False
        self.awaiting_inn = False


async def set_commands(app):
    commands = [
        ("start", "Начать работу"),
        ("new", "Новая декларация"),
        ("status", "Мой статус"),
        ("help", "Помощь"),
    ]
    await app.bot.set_my_commands(commands)


async def notify_admin(context, text):
    for admin_id in ADMIN_IDS:
        try:
            await context.bot.send_message(chat_id=admin_id, text=text)
        except Exception as e:
            print(f"Не удалось отправить уведомление админу {admin_id}: {e}")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user = update.effective_user
    
    update_user_data(
        user_id,
        username=user.username,
        first_name=user.first_name,
        last_name=user.last_name
    )
    
    user_sessions[user_id] = UserSession(user_id)
    
    user_info = f"@{user.username}" if user.username else f"{user.first_name} {user.last_name or ''}"
    await notify_admin(context, f"Новый пользователь!\n\n👤 {user_info}\n🆔 ID: {user_id}")
    
    await update.message.reply_text(
        "🤖 *Бот для подготовки отчетности ИП на УСН (Доходы 6%)*\n\n"
        "1️⃣ Загрузите выписки с расчетных счетов (Excel)\n"
        "2️⃣ Загрузите выписку с ЕНС (CSV)\n\n"
        "📌 *Сроки за 2025 год:*\n"
        "• Декларацию сдать до 27 апреля 2026\n"
        "• Налог уплатить до 28 апреля 2026\n\n"
        f"💰 *Тарифы:*\n"
        f"• Демо: 3 попытки (только Титул + Раздел 1.1)\n"
        f"• Полная версия: {SUBSCRIPTION_PRICE}₽/мес (все разделы + XML)",
        parse_mode="Markdown"
    )


async def new_declaration(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id in user_sessions:
        user_sessions[user_id].reset()
    
    await update.message.reply_text(
        "Начинаем новую декларацию!\n\n"
        "1️⃣ Загрузите выписки с расчетных счетов (Excel)\n"
        "2️⃣ Загрузите выписку с ЕНС (CSV)"
    )


async def my_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if can_use_full_version(user_id):
        if is_admin(user_id):
            await update.message.reply_text("✅ *Ваш статус:* администратор (полный доступ)", parse_mode="Markdown")
        else:
            user_data = get_user_data(user_id)
            until = datetime.fromisoformat(user_data["subscription_until"])
            days_left = (until - datetime.now()).days
            await update.message.reply_text(
                f"✅ *Ваш статус:* активная подписка\n\n"
                f"📅 Действует до: {until.strftime('%d.%m.%Y')}\n"
                f"📊 Осталось дней: {days_left}",
                parse_mode="Markdown"
            )
    else:
        attempts_left = get_demo_attempts_left(user_id)
        await update.message.reply_text(
            f"⚠️ *Ваш статус:* демо-доступ\n\n"
            f"📊 Осталось попыток: {attempts_left} из 3\n\n"
            f"💰 *Полная версия:* {SUBSCRIPTION_PRICE}₽/мес\n"
            f"✅ Все разделы декларации\n"
            f"✅ XML для загрузки в ЛК ФНС\n\n"
            f"📞 Для оплаты свяжитесь с администратором: @{ADMIN_USERNAME}",
            parse_mode="Markdown"
        )


async def contact_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user = update.effective_user
    
    username_info = f"@{user.username}" if user.username else "Username не указан"
    
    await notify_admin(
        context,
        f"Запрос связи от пользователя\n\n"
        f"👤 {user.first_name} {user.last_name or ''}\n"
        f"🆔 ID: {user_id}\n"
        f"📱 Username: {username_info}"
    )
    
    await update.message.reply_text(
        f"📞 *Запрос отправлен!*\n\n"
        f"Администратор свяжется с вами в ближайшее время.\n\n"
        f"📱 Для связи: @{ADMIN_USERNAME}",
        parse_mode="Markdown"
    )


async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not is_admin(user_id):
        await update.message.reply_text("⛔ У вас нет доступа")
        return
    
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Список пользователей", callback_data="admin_users")],
        [InlineKeyboardButton("Дать доступ", callback_data="admin_add_access")],
        [InlineKeyboardButton("Статистика", callback_data="admin_stats")],
        [InlineKeyboardButton("Рассылка", callback_data="admin_broadcast")],
    ])
    
    await update.message.reply_text(
        "⚙️ *Админ панель*\n\nВыберите действие:",
        parse_mode="Markdown",
        reply_markup=keyboard
    )


async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    
    if not is_admin(user_id):
        await query.edit_message_text("⛔ У вас нет доступа")
        return
    
    if query.data == "admin_users":
        users = load_users()
        paid = 0
        demo = 0
        for uid, data in users.items():
            if data.get("subscription_until"):
                try:
                    until = datetime.fromisoformat(data["subscription_until"])
                    if datetime.now() < until:
                        paid += 1
                    else:
                        demo += 1
                except:
                    demo += 1
            else:
                demo += 1
        
        text = f"Статистика пользователей\n\n✅ Платных: {paid}\n⚠️ Демо: {demo}\n📊 Всего: {len(users)}"
        await query.edit_message_text(text)
        
        back_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("◀️ Назад", callback_data="admin_back")]
        ])
        await query.edit_message_reply_markup(reply_markup=back_keyboard)
    
    elif query.data == "admin_add_access":
        await query.edit_message_text(
            "Добавление доступа\n\n"
            "Введите команду:\n"
            "/add <user_id> <days>\n\n"
            "Пример: /add 123456789 30"
        )
    
    elif query.data == "admin_stats":
        users = load_users()
        total_ops = 0
        for uid in users:
            if uid in user_sessions:
                total_ops += len(user_sessions[uid].bank_operations)
        
        text = f"Статистика\n\n👥 Пользователей: {len(users)}\n💳 Операций обработано: {total_ops}\n💰 Цена подписки: {SUBSCRIPTION_PRICE}₽/мес"
        await query.edit_message_text(text)
        
        back_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("◀️ Назад", callback_data="admin_back")]
        ])
        await query.edit_message_reply_markup(reply_markup=back_keyboard)
    
    elif query.data == "admin_broadcast":
        context.user_data['broadcast_mode'] = True
        await query.edit_message_text(
            "Рассылка\n\nВведите сообщение для рассылки всем пользователям:"
        )
    
    elif query.data == "admin_back":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("Список пользователей", callback_data="admin_users")],
            [InlineKeyboardButton("Дать доступ", callback_data="admin_add_access")],
            [InlineKeyboardButton("Статистика", callback_data="admin_stats")],
            [InlineKeyboardButton("Рассылка", callback_data="admin_broadcast")],
        ])
        await query.edit_message_text(
            "⚙️ Админ панель\n\nВыберите действие:",
            reply_markup=keyboard
        )


async def add_subscription(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if not is_admin(user_id):
        await update.message.reply_text("⛔ У вас нет доступа")
        return
    
    try:
        target_user_id = int(context.args[0])
        days = int(context.args[1])
    except (IndexError, ValueError):
        await update.message.reply_text(
            "❌ Использование: /add <user_id> <days>\n"
            "Пример: /add 123456789 30"
        )
        return
    
    until = datetime.now() + timedelta(days=days)
    update_user_data(target_user_id, subscription_until=until.isoformat())
    
    await update.message.reply_text(
        f"✅ Пользователю {target_user_id} добавлен доступ на {days} дней\n"
        f"📅 Действует до: {until.strftime('%d.%m.%Y')}"
    )
    
    try:
        await context.bot.send_message(
            chat_id=target_user_id,
            text=f"🎉 Вам открыт полный доступ до {until.strftime('%d.%m.%Y')}!\n\n"
                 f"Теперь вы можете получить полную декларацию с XML.\n"
                 f"Отправьте /new и загрузите выписки."
        )
    except:
        pass


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
            bank_name = detect_bank_name(filename)
            operations, inn, fio, accounts = parse_bank_statement(tmp_path)
            
            if operations:
                session.add_bank_operations(operations, bank_name, inn, fio, accounts)
                total = sum(op['amount'] for op in operations)
                total_all = sum(op['amount'] for op in session.bank_operations)
                
                await update.message.reply_text(
                    f"✅ {bank_name}: {len(operations)} операций, {total:,.2f} ₽"
                )
                
                if not session.ens_loaded:
                    await update.message.reply_text(
                        f"📊 *Итог по банковским выпискам:*\n"
                        f"• Количество файлов: {len(session.bank_files)}\n"
                        f"• Доход за 2025: {total_all:,.2f} ₽\n\n"
                        f"📎 *Теперь пришлите выписку из ЕНС (CSV)*\n"
                        f"Ее можно скачать в личном кабинете налогоплательщика",
                        parse_mode="Markdown"
                    )
            else:
                await update.message.reply_text(f"⚠️ В выписке из {bank_name} не найдено доходов")
        
        elif filename.endswith('.csv'):
            await update.message.reply_text("📥 Обрабатываю выписку ЕНС...")
            ens_data = parse_ens_statement(tmp_path)
            session.set_ens_data(ens_data)
            
            total_income = sum(op['amount'] for op in session.bank_operations)
            paid_in_2025 = any(d.year == 2025 for d in ens_data.get('insurance_paid_dates', []))
            
            msg = f"✅ Выписка ЕНС обработана!\n\n"
            msg += f"📊 *Итог по банковским выпискам:*\n"
            msg += f"• Количество файлов: {len(session.bank_files)}\n"
            msg += f"• Доход за 2025: {total_income:,.2f} ₽\n\n"
            msg += f"📌 *Данные из ЕНС:*\n"
            msg += f"• Начислено взносов: {ens_data['insurance_accrued']:,.2f} ₽\n"
            msg += f"• Уплачено взносов: {ens_data['insurance_paid']:,.2f} ₽\n"
            msg += f"• Уплачено в 2025: {'Да' if paid_in_2025 else 'Нет'}\n"
            msg += f"• Авансов по УСН: {len(ens_data['usn_payments'])}\n"
            msg += f"• ОКТМО: {session.oktmo if session.oktmo else 'не найден'}\n\n"
            
            await update.message.reply_text(msg, parse_mode="Markdown")
            
            await ask_missing_data(update, session)
        
        else:
            await update.message.reply_text("❌ Поддерживаются .xlsx, .xls, .csv")
    
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")
    
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


async def ask_missing_data(update: Update, session):
    if not session.phone:
        session.awaiting_phone = True
        await update.message.reply_text(
            "📞 *Для заполнения декларации нужен контактный телефон*\n"
            "Например: *81234567890*",
            parse_mode="Markdown"
        )
        return
    
    if not session.oktmo or session.oktmo == "36701320":
        session.awaiting_oktmo = True
        await update.message.reply_text(
            "📝 *Укажите код ОКТМО*\n"
            "Его можно найти в выписке ЕНС или на сайте ФНС\n\n"
            "Например: *36701000*",
            parse_mode="Markdown"
        )
        return
    
    if not session.fio:
        session.awaiting_fio = True
        await update.message.reply_text(
            "📝 *Укажите ФИО полностью*\n"
            "Например: *Иванов Иван Иванович*",
            parse_mode="Markdown"
        )
        return
    
    await update.message.reply_text("✅ Все данные получены! Формирую декларацию...")
    await generate_and_send_report(update, session)


async def generate_and_send_report(update: Update, session):
    user_id = session.user_id
    
    try:
        all_ops = []
        for op in session.bank_operations:
            if isinstance(op, dict):
                all_ops.append(op)
            elif isinstance(op, list):
                all_ops.extend(op)
        all_ops.sort(key=lambda x: x['date'])
        
        decl_template = os.path.join(TEMPLATES_DIR, "Declaration_template.xlsx")
        
        if not os.path.exists(decl_template):
            await update.message.reply_text(f"❌ Шаблон декларации не найден")
            return
        
        inn = session.inn if session.inn else ""
        fio = session.fio if session.fio else ""
        oktmo = session.oktmo if session.oktmo else ""
        ip_accounts = session.ip_accounts if session.ip_accounts else []
        phone = session.phone if session.phone else ""
        
        is_full = can_use_full_version(user_id)
        
        if not is_full:
            attempts_left = use_demo_attempt(user_id)
            if attempts_left < 0:
                await update.message.reply_text(
                    f"❌ *Лимит демо-попыток исчерпан!*\n\n"
                    f"💰 Приобретите полную версию за {SUBSCRIPTION_PRICE}₽/мес\n"
                    f"📞 Свяжитесь с администратором: @{ADMIN_USERNAME}",
                    parse_mode="Markdown"
                )
                return
        
        decl_excel, decl_xml, total_income, tax_payable = generate_report(
            all_ops, session.ens_data, OUTPUT_DIR, user_id,
            decl_template, inn, fio, oktmo, ip_accounts, phone,
            is_full_version=is_full
        )
        
        if is_full:
            await update.message.reply_text(
                f"✅ *Декларация готова!*\n\n"
                f"📊 Доход за 2025: {total_income:,.2f} ₽\n"
                f"💰 Налог к уплате: {tax_payable:,.2f} ₽\n\n"
                f"⚠️ *Проверьте правильность указанных данных (желтые ячейки)*\n\n"
                f"📌 Сдать декларацию до 27 апреля 2026\n"
                f"📌 Уплатить налог до 28 апреля 2026",
                parse_mode="Markdown"
            )
            
            with open(decl_excel, 'rb') as f:
                await update.message.reply_document(f, filename="Декларация_УСН_2025.xlsx", caption="📝 Полная декларация")
            
            if decl_xml and os.path.exists(decl_xml):
                with open(decl_xml, 'rb') as f:
                    await update.message.reply_document(f, filename="declaration_usn_2025.xml", caption="📎 XML для ЛК ФНС")
        else:
            remaining = get_demo_attempts_left(user_id)
            await update.message.reply_text(
                f"⚠️ *Демо-версия декларации*\n\n"
                f"📊 Доход за 2025: {total_income:,.2f} ₽\n"
                f"🔒 *Сумма налога скрыта (limited)*\n\n"
                f"📌 Чтобы получить полную декларацию с XML:\n"
                f"💰 {SUBSCRIPTION_PRICE}₽/мес\n"
                f"📞 Свяжитесь с администратором: @{ADMIN_USERNAME}\n\n"
                f"ℹ️ Осталось попыток: {remaining} из 3",
                parse_mode="Markdown"
            )
            
            with open(decl_excel, 'rb') as f:
                await update.message.reply_document(f, filename="Декларация_УСН_2025_ДЕМО.xlsx", caption="📝 Демо-версия")
    
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")
        import traceback
        traceback.print_exc()


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    
    if context.user_data.get('broadcast_mode') and is_admin(user_id):
        users = load_users()
        success = 0
        fail = 0
        
        await update.message.reply_text("📢 Начинаю рассылку...")
        
        for uid in users.keys():
            try:
                await context.bot.send_message(chat_id=int(uid), text=text)
                success += 1
            except:
                fail += 1
        
        await update.message.reply_text(f"✅ Рассылка завершена\n📨 Доставлено: {success}\n❌ Ошибок: {fail}")
        context.user_data['broadcast_mode'] = False
        return
    
    if user_id not in user_sessions:
        user_sessions[user_id] = UserSession(user_id)
    
    session = user_sessions[user_id]
    
    if session.awaiting_phone:
        phone_digits = ''.join(ch for ch in text if ch.isdigit())
        if phone_digits:
            session.phone = phone_digits
            session.awaiting_phone = False
            await update.message.reply_text(f"✅ Телефон сохранен: {phone_digits}")
            await ask_missing_data(update, session)
        else:
            await update.message.reply_text("❌ Введите номер телефона цифрами")
        return
    
    if session.awaiting_oktmo:
        oktmo_digits = ''.join(ch for ch in text if ch.isdigit())
        if len(oktmo_digits) >= 8:
            session.oktmo = oktmo_digits[:8]
            session.awaiting_oktmo = False
            await update.message.reply_text(f"✅ ОКТМО сохранен: {session.oktmo}")
            await ask_missing_data(update, session)
        else:
            await update.message.reply_text("❌ ОКТМО должен содержать 8 цифр")
        return
    
    if session.awaiting_fio:
        if len(text.split()) >= 2:
            session.fio = text
            session.awaiting_fio = False
            await update.message.reply_text(f"✅ ФИО сохранено: {text}")
            await ask_missing_data(update, session)
        else:
            await update.message.reply_text("❌ Введите полное ФИО (Фамилия Имя Отчество)")
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
        "/start — начать работу\n"
        "/new — начать новую декларацию\n"
        "/status — проверить статус подписки\n"
        "/reset — сбросить текущую сессию\n"
        "/help — показать это сообщение\n\n"
        f"📞 По вопросам оплаты: @{ADMIN_USERNAME}",
        parse_mode="Markdown"
    )


def main():
    if not BOT_TOKEN:
        print("❌ BOT_TOKEN не задан в .env")
        sys.exit(1)
    
    app = Application.builder().token(BOT_TOKEN).build()
    
    app.post_init = set_commands
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("new", new_declaration))
    app.add_handler(CommandHandler("status", my_status))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("add", add_subscription))
    
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(CallbackQueryHandler(admin_callback))
    
    print("🤖 Бот запущен...")
    app.run_polling()


if __name__ == "__main__":
    main()
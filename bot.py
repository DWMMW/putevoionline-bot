from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from openpyxl import load_workbook
from datetime import datetime, timedelta
import os
import pythoncom
import win32com.client
import json
import asyncio
import requests

# Конфигурация
TOKEN = "7925214932:AAEMghAZxNU1RNPhFtSzqDA3n1rSbvgSZuM"  # Ваш токен бота
ENOT_API_KEY = "b0f8c01244a772a93f8afb77b86ce85b47a1d1e3"  # Ваш API-ключ ENOT.io
ADMIN_ID = 6926880256  # Ваш Telegram ID
bot = Bot(token=TOKEN)
dp = Dispatcher()

# Константы
FREE_FILE_LIMIT = 20
FILE_COUNTS_PATH = "file_counts.json"
PAYMENTS_PATH = "payments.json"
REPORTS_PATH = "user_reports.json"
USER_DATA_PATH = "user_data.json"

# Базовый список полей
base_fields = [
    ("Ввести номер телефона", "phone_number"),
    ("Ввести данные ФИО водителя", "driver_name"),
    ("Ввести данные ИНН", "inn"),
    ("Указать ИНН если Вы самозанятый или указать Ваше ИП", "self_employed_inn"),
    ("СНИЛС", "snils"),
    ("Ввести номер т/с", "vehicle_number"),
    ("Ввести марку, модель легкового автомобиля", "vehicle_model"),
    ("Адрес организации", "organization_address"),
    ("Водительское удостоверение", "license_number"),
    ("Срок действия ВУ с", "license_from"),
    ("Срок действия ВУ по", "license_to"),
    ("Лицензия на перевозку, регистрационный №", "transport_license"),
    ("С какой даты действует лицензия на перевозку", "license_date"),
    ("Показания одометра", "odometer")
]

ogrn_field = ("Введите ОГРН", "ogrn")

# Хранилища
user_data = {}
file_counts = {}
phone_to_user_mapping = {}
paid_files = {}
user_reports = {}

# Функции для работы с файлами
def load_file_counts():
    global file_counts
    try:
        if os.path.exists(FILE_COUNTS_PATH):
            with open(FILE_COUNTS_PATH, "r") as f:
                file_counts = json.load(f)
                file_counts = {int(k): v for k, v in file_counts.items()}
    except (json.JSONDecodeError, FileNotFoundError):
        file_counts = {}

def save_file_counts():
    with open(FILE_COUNTS_PATH, "w") as f:
        json.dump(file_counts, f)

def load_paid_files():
    global paid_files
    try:
        if os.path.exists(PAYMENTS_PATH):
            with open(PAYMENTS_PATH, "r") as f:
                paid_files = json.load(f)
                paid_files = {int(k): v for k, v in paid_files.items()}
    except (json.JSONDecodeError, FileNotFoundError):
        paid_files = {}

def save_paid_files():
    with open(PAYMENTS_PATH, "w") as f:
        json.dump(paid_files, f)

def load_user_reports():
    global user_reports
    try:
        if os.path.exists(REPORTS_PATH):
            with open(REPORTS_PATH, "r") as f:
                user_reports = json.load(f)
                user_reports = {int(k): v for k, v in user_reports.items()}
    except (json.JSONDecodeError, FileNotFoundError):
        user_reports = {}

def save_user_reports():
    with open(REPORTS_PATH, "w") as f:
        json.dump(user_reports, f)

def load_user_data():
    global user_data, phone_to_user_mapping
    try:
        if os.path.exists(USER_DATA_PATH):
            with open(USER_DATA_PATH, "r") as f:
                loaded_data = json.load(f)
                user_data = {int(k): v for k, v in loaded_data["user_data"].items()}
                phone_to_user_mapping = loaded_data["phone_to_user_mapping"]
                phone_to_user_mapping = {k: int(v) for k, v in phone_to_user_mapping.items()}
                for uid in user_data:
                    user_data[uid]["processing_pdf"] = False
    except (json.JSONDecodeError, FileNotFoundError):
        user_data = {}
        phone_to_user_mapping = {}

def save_user_data():
    with open(USER_DATA_PATH, "w") as f:
        json.dump({"user_data": user_data, "phone_to_user_mapping": phone_to_user_mapping}, f)

# Загрузка данных при старте
load_file_counts()
load_paid_files()
load_user_reports()
load_user_data()

# Сопоставление ячеек
cell_mapping = {
    "phone_number": (8, 26), "inn": (4, 4), "self_employed_inn": (8, 6),
    "driver_name": (12, 7), "vehicle_number": (11, 12), "vehicle_model": (10, 15),
    "snils": (12, 22), "license_number": (13, 15), "license_from": (13, 26), "license_to": (13, 31),
    "transport_license": (14, 17), "license_date": (14, 26), "organization_address": (9, 8),
    "odometer": (20, 12), "ogrn": (3, 4),
    "departure_date": (18, 12), "tech_control_date": (18, 35),
    "departure_time_minus_11": (19, 12), "tech_control_time_minus_22": (20, 35),
    "medical_check_time_minus_33": (51, 12),
    "document_date": (6, 14), "next_day_date": (6, 26), "file_number": (4, 17),
    "departure_hour": (41, 23), "departure_minute": (41, 25),
    "vehicle_number_duplicate_1": (11, 35), "vehicle_number_duplicate_2": (12, 35),
    "medical_check_date": (50, 12)
}

# Клавиатуры
start_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Получить путевой лист")]],
    resize_keyboard=True,
    one_time_keyboard=True
)

finish_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Получить путевой лист")],
        [KeyboardButton(text="Изменить данные")]
    ],
    resize_keyboard=True
)

input_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Назад"), KeyboardButton(text="Отмена")]
    ],
    resize_keyboard=True
)

payment_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="60 файлов - 600р (50% скидка)")],
        [KeyboardButton(text="28 файлов - 350р (50% скидка)")],
        [KeyboardButton(text="14 файлов - 420р (50% скидка)")],
        [KeyboardButton(text="2 файла - 70р (50% скидка)")],
        [KeyboardButton(text="Отмена")]
    ],
    resize_keyboard=True
)

confirm_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Подтвердить")],
        [KeyboardButton(text="Отмена")]
    ],
    resize_keyboard=True
)

choice_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Изменить показания одометра")],
        [KeyboardButton(text="Изменить все данные")]
    ],
    resize_keyboard=True
)

# Обработчики
@dp.message(Command("start"))
async def start_command(message: types.Message):
    if message.from_user is None:
        return
    welcome_text = (
        "Добрый день! Вас приветствует Путевой Онлайн. "
        "Все Ваши данные шифруются, не сохраняются и не передаются третьим лицам.\n\n"
        "Нажмите 'Получить путевой лист', чтобы приступить."
    )
    await message.answer(welcome_text, reply_markup=start_keyboard)

@dp.message(lambda message: message.from_user.id not in user_data and message.text != "Получить путевой лист")
async def start_new_user(message: types.Message):
    if message.from_user is None:
        return
    welcome_text = (
        "Добрый день! Вас приветствует Путевой Онлайн. "
        "Все Ваши данные шифруются, не сохраняются и не передаются третьим лицам.\n\n"
        "Нажмите 'Получить путевой лист', чтобы приступить."
    )
    await message.answer(welcome_text, reply_markup=start_keyboard)

@dp.message(lambda message: message.text == "Получить путевой лист")
async def handle_start(message: types.Message):
    user_id = message.from_user.id
    if user_id not in file_counts:
        file_counts[user_id] = 0
    
    if user_id in user_data and user_data[user_id].get("processing_pdf", False):
        await message.answer("Идет создание предыдущего файла, пожалуйста, подождите.")
        return
    
    phone = next((p for p, uid in phone_to_user_mapping.items() if uid == user_id), None)
    if phone and user_id in user_data and "phone_number" in user_data[user_id]["data"]:
        required_fields = [field[1] for field in base_fields if field[1] != "odometer"]
        has_all_data = all(field in user_data[user_id]["data"] for field in required_fields)
        if has_all_data:
            user_data[user_id]["step"] = len(user_data[user_id]["full_fields"]) - 1
            user_data[user_id]["processing_pdf"] = False
            save_user_data()
            await message.answer("Ваши данные найдены. Что вы хотите сделать?", reply_markup=choice_keyboard)
        else:
            user_data[user_id] = {"step": 0, "data": {"phone_number": phone}, "full_fields": base_fields.copy(), "processing_pdf": False}
            await ask_next_field(message)
    else:
        user_data[user_id] = {"step": 0, "data": {}, "full_fields": base_fields.copy(), "processing_pdf": False}
        await ask_next_field(message)

@dp.message(lambda message: message.text == "Изменить показания одометра")
async def edit_odometer(message: types.Message):
    user_id = message.from_user.id
    if user_id in user_data:
        user_data[user_id]["step"] = len(user_data[user_id]["full_fields"]) - 1
        await message.answer("Введите новые показания одометра", reply_markup=confirm_keyboard)

@dp.message(lambda message: message.text in [
    "60 файлов - 600р (50% скидка)", "28 файлов - 350р (50% скидка)",
    "14 файлов - 420р (50% скидка)", "2 файла - 70р (50% скидка)"
])
async def process_payment_selection(message: types.Message):
    user_id = message.from_user.id
    packages = {
        "60 файлов - 600р (50% скидка)": (60, 600),
        "28 файлов - 350р (50% скидка)": (28, 350),
        "14 файлов - 420р (50% скидка)": (14, 420),
        "2 файла - 70р (50% скидка)": (2, 70)
    }
    package_text = message.text
    files, amount = packages[package_text]

    # Создание платежа через ENOT.io API
    order_id = f"{user_id}_{int(datetime.now().timestamp())}"  # Уникальный ID заказа
    headers = {"Content-Type": "application/json"}
    data = {
        "api_key": ENOT_API_KEY,
        "amount": amount,
        "currency": "RUB",
        "order_id": order_id,
        "description": f"Покупка {files} файлов для Путевого Онлайн",
        "email": "anonymous@temp-mail.org"  # Для анонимности
    }
    response = requests.post("https://enot.io/api/v1/invoice/create", headers=headers, json=data)
    
    if response.status_code == 200:
        payment_data = response.json()
        payment_url = payment_data["url"]
        await message.answer(
            f"Оплатите {amount}р по ссылке для получения {files} файлов:\n{payment_url}",
            reply_markup=ReplyKeyboardRemove()
        )
        user_data[user_id]["pending_payment"] = {"files": files, "amount": amount, "order_id": order_id}
        save_user_data()
    else:
        await message.answer(f"Ошибка при создании счета: {response.status_code} - {response.text}. Попробуйте позже.", reply_markup=finish_keyboard)

@dp.message(lambda message: message.text == "Подтвердить")
async def confirm_odometer(message: types.Message):
    user_id = message.from_user.id
    if user_id not in user_data or user_data[user_id]["step"] < len(user_data[user_id]["full_fields"]):
        await message.answer("Сначала введите показания одометра.", reply_markup=confirm_keyboard)
        return
    
    if "odometer" not in user_data[user_id]["data"]:
        await message.answer("Сначала введите показания одометра.", reply_markup=confirm_keyboard)
        return
    
    if user_data[user_id].get("processing_pdf", False):
        await message.answer("Идет создание предыдущего файла, пожалуйста, подождите.")
        return
    
    user_data[user_id]["processing_pdf"] = True
    try:
        await message.answer("Создание файла, подождите...")
        await create_pdf(message)
    except Exception as e:
        await message.answer(f"Ошибка при создании файла: {str(e)}")
    finally:
        user_data[user_id]["processing_pdf"] = False
        save_user_data()

@dp.message(lambda message: message.text == "Изменить все данные")
async def edit_data(message: types.Message):
    user_id = message.from_user.id
    if user_id in user_data:
        phone = user_data[user_id]["data"].get("phone_number", "")
        user_data[user_id] = {"step": 1, "data": {"phone_number": phone}, "full_fields": base_fields.copy(), "processing_pdf": False}
        save_user_data()
        await message.answer("Введите данные ФИО водителя", reply_markup=input_keyboard)
    else:
        await message.answer("Сначала создайте хотя бы один путевой лист.", reply_markup=start_keyboard)

@dp.message(Command("report"))
async def send_report(message: types.Message):
    user_id = message.from_user.id
    if user_id != ADMIN_ID:
        await message.answer("Эта команда доступна только администратору.")
        return
    
    report_text = "Отчет по пользователям:\n\n"
    for uid in user_reports:
        report = user_reports[uid]
        phone = report.get("phone_number", "Не указан")
        free_files = min(file_counts.get(uid, 0), FREE_FILE_LIMIT)
        paid_files_count = max(file_counts.get(uid, 0) - FREE_FILE_LIMIT, 0)
        payments = report.get("payments", [])
        total_paid = sum(p["amount"] for p in payments)
        report_text += (
            f"ID: {uid}\n"
            f"Телефон: {phone}\n"
            f"Бесплатные файлы: {free_files}\n"
            f"Платные файлы: {paid_files_count}\n"
            f"Всего оплат: {len(payments)}\n"
            f"Сумма оплат: {total_paid}р\n"
            f"История оплат: {payments}\n\n"
        )
    
    if not user_reports:
        report_text = "Нет данных о пользователях."
    
    await message.answer(report_text)

async def ask_next_field(message: types.Message):
    user_id = message.from_user.id
    step = user_data[user_id]["step"]
    full_fields = user_data[user_id]["full_fields"]
    
    if step < len(full_fields):
        await message.answer(f"Введите {full_fields[step][0]}", reply_markup=input_keyboard)
    else:
        await message.answer("Введите новые показания одометра", reply_markup=confirm_keyboard)

@dp.message(lambda message: message.text == "Назад")
async def go_back(message: types.Message):
    user_id = message.from_user.id
    if user_id in user_data and user_data[user_id]["step"] > 1:
        user_data[user_id]["step"] -= 1
        step = user_data[user_id]["step"]
        full_fields = user_data[user_id]["full_fields"]
        await message.answer(f"Введите {full_fields[step][0]}", reply_markup=input_keyboard)
    else:
        keyboard = finish_keyboard if user_id in file_counts and file_counts[user_id] > 0 else start_keyboard
        await message.answer("Вы вернулись к началу.", reply_markup=keyboard)

@dp.message(lambda message: message.text == "Отмена")
async def cancel_input(message: types.Message):
    user_id = message.from_user.id
    if user_id in user_data and user_data[user_id]["step"] > 0:
        user_data[user_id]["step"] = 0
        keyboard = finish_keyboard if user_id in file_counts and file_counts[user_id] > 0 else start_keyboard
        await message.answer("Ввод данных отменен.", reply_markup=keyboard)
    else:
        await message.answer("Нечего отменять.", reply_markup=start_keyboard)

@dp.message()
async def save_data(message: types.Message):
    if message.from_user is None:
        return
    user_id = message.from_user.id
    if user_id not in user_data:
        await start_new_user(message)
        return
    
    step = user_data[user_id]["step"]
    full_fields = user_data[user_id]["full_fields"]
    
    if step < len(full_fields):
        if full_fields[step][1] == "phone_number":
            phone = message.text
            if phone in phone_to_user_mapping:
                if phone_to_user_mapping[phone] != user_id:
                    await message.answer("Этот номер телефона уже привязан к другому аккаунту Telegram. Введите другой номер.", reply_markup=input_keyboard)
                    return
                else:
                    user_data[user_id]["step"] = len(full_fields) - 1
                    await message.answer("Ваши данные найдены. Что вы хотите сделать?", reply_markup=choice_keyboard)
                    save_user_data()
                    return
            else:
                phone_to_user_mapping[phone] = user_id
                if user_id not in user_reports:
                    user_reports[user_id] = {"phone_number": phone, "payments": []}
                else:
                    user_reports[user_id]["phone_number"] = phone
                save_user_reports()
        
        elif full_fields[step][1] == "self_employed_inn":
            input_text = message.text.strip()
            if input_text.upper().startswith("ИП"):
                self_employed_index = [f[1] for f in full_fields].index("self_employed_inn")
                full_fields.insert(self_employed_index + 1, ogrn_field)
                user_data[user_id]["full_fields"] = full_fields
            elif not input_text.isdigit():
                await message.answer("Введите либо числовой ИНН, либо данные, начинающиеся с 'ИП'.", reply_markup=input_keyboard)
                return
        
        user_data[user_id]["data"][full_fields[step][1]] = message.text
        user_data[user_id]["step"] += 1
        save_user_data()
        await ask_next_field(message)
    else:
        user_data[user_id]["data"]["odometer"] = message.text
        await message.answer(f"Вы ввели одометр: {message.text}. Нажмите 'Подтвердить' для создания файла.", reply_markup=confirm_keyboard)

async def create_pdf(message: types.Message):
    user_id = message.from_user.id
    pdf_path = f"user_{user_id}_document.pdf"
    
    current_file_count = file_counts.get(user_id, 0)
    paid_files_left = paid_files.get(user_id, 0)
    
    if current_file_count >= FREE_FILE_LIMIT and paid_files_left <= 0:
        await message.answer("Вы использовали все бесплатные 20 файлов. Выберите пакет для оплаты:", reply_markup=payment_keyboard)
        user_data[user_id]["processing_pdf"] = False
        save_user_data()
        return
    
    file_number = current_file_count + 1
    
    current_datetime = datetime.now()
    mo_datetime = current_datetime - timedelta(minutes=33)
    to_datetime = current_datetime - timedelta(minutes=22)
    expiry_datetime = mo_datetime + timedelta(hours=12)
    
    await message.answer(f"Текущий путевой лист: {file_number} от {current_datetime.strftime('%d.%m.%Y %H:%M:%S')}", reply_markup=ReplyKeyboardRemove())
    await message.answer(f"Дата и время МО: {mo_datetime.strftime('%d.%m.%Y %H:%M:%S')}")
    await message.answer(f"Дата и время ТО: {to_datetime.strftime('%d.%m.%Y %H:%M:%S')}")
    await message.answer(f"Срок действия путевого листа до: {expiry_datetime.strftime('%d.%m.%Y %H:%M:%S')}")
    
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.abspath("template.xlsx"))
        ws = wb.ActiveSheet
        
        user_data[user_id]["data"].update({
            "departure_date": current_datetime.strftime("%d.%m.%Y"),
            "tech_control_date": current_datetime.strftime("%d.%m.%Y"),
            "medical_check_date": current_datetime.strftime("%d.%m.%Y"),
            "departure_time_minus_11": (current_datetime - timedelta(minutes=11)).strftime("%H:%M"),
            "tech_control_time_minus_22": to_datetime.strftime("%H:%M"),
            "medical_check_time_minus_33": mo_datetime.strftime("%H:%M"),
            "document_date": current_datetime.strftime("%d.%m.%Y"),
            "next_day_date": (current_datetime + timedelta(days=1)).strftime("%d.%m.%Y"),
            "file_number": file_number,
            "departure_hour": (current_datetime - timedelta(minutes=11)).strftime("%H"),
            "departure_minute": (current_datetime - timedelta(minutes=11)).strftime("%M"),
            "vehicle_number_duplicate_1": user_data[user_id]["data"].get("vehicle_number", ""),
            "vehicle_number_duplicate_2": user_data[user_id]["data"].get("vehicle_number", "")
        })
        
        for field, (row, col) in cell_mapping.items():
            if field in user_data[user_id]["data"]:
                ws.Cells(row, col).Value = user_data[user_id]["data"][field]
        
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        wb.Close(SaveChanges=False)
        excel.Quit()
        
        await message.answer_document(types.FSInputFile(pdf_path))
        file_counts[user_id] = file_number
        if file_number > FREE_FILE_LIMIT:
            paid_files[user_id] = paid_files.get(user_id, 0) - 1
        save_file_counts()
        save_paid_files()
        
        if user_id not in user_reports:
            user_reports[user_id] = {"phone_number": user_data[user_id]["data"].get("phone_number", "Не указан"), "payments": []}
        save_user_reports()
        
        await message.answer("Спасибо, что воспользовались сервисом Путевой Онлайн!", reply_markup=finish_keyboard)
    except Exception as e:
        await message.answer(f"Ошибка при создании PDF: {str(e)}")
        raise
    finally:
        user_data[user_id]["processing_pdf"] = False
        save_user_data()
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

if __name__ == "__main__":
    import asyncio
    async def main():
        await dp.start_polling(bot)
    asyncio.run(main())
from flask import Flask, request
import json
import hmac
import hashlib
from datetime import datetime

app = Flask(__name__)

# Секретный ключ вебхука ENOT.io
ENOT_SECRET_KEY = "036819f02f6d855429c189d91683d1c6ac9c146b"

# Функция для проверки подписи
def verify_signature(payload, signature):
    computed_signature = hmac.new(
        ENOT_SECRET_KEY.encode('utf-8'),
        json.dumps(payload, separators=(',', ':')).encode('utf-8'),
        hashlib.sha256
    ).hexdigest()
    return hmac.compare_digest(computed_signature, signature)

@app.route('/enot_callback', methods=['POST'])
def enot_callback():
    # Получаем данные и подпись
    data = request.json
    signature = request.headers.get('X-Signature', '')

    # Проверяем подпись
    if not verify_signature(data, signature):
        return "Invalid signature", 403

    order_id = data.get("order_id")
    status = data.get("status")
    user_id = int(order_id.split("_")[0])  # Извлекаем user_id

    if status == "success":
        # Загружаем данные
        with open("user_data.json", "r") as f:
            user_data = json.load(f)["user_data"]
        with open("payments.json", "r") as f:
            paid_files = json.load(f)
        with open("user_reports.json", "r") as f:
            user_reports = json.load(f)

        # Проверяем ожидаемый платеж
        pending = user_data.get(str(user_id), {}).get("pending_payment", {})
        if pending and pending["order_id"] == order_id:
            files = pending["files"]
            amount = pending["amount"]

            # Начисляем файлы
            paid_files[str(user_id)] = paid_files.get(str(user_id), 0) + files
            if str(user_id) not in user_reports:
                user_reports[str(user_id)] = {"phone_number": user_data[str(user_id)]["data"].get("phone_number", "Не указан"), "payments": []}
            user_reports[str(user_id)]["payments"].append({
                "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "amount": amount,
                "files": files
            })

            # Сохраняем изменения
            with open("payments.json", "w") as f:
                json.dump(paid_files, f)
            with open("user_reports.json", "w") as f:
                json.dump(user_reports, f)

            # Уведомляем пользователя
            loop = asyncio.get_event_loop()
            loop.run_until_complete(bot.send_message(user_id, f"Оплата прошла успешно! Вам начислено {files} файлов."))

    return "OK", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
import openpyxl
from openpyxl.styles import Font, Alignment
from io import BytesIO
from telegram import Update
from telegram.ext import Application, MessageHandler, filters
import json

TOKEN = "8726758237:AAH66M1yXIAkb2ksNn-sQuGMdy-a0y9PGnc"

async def handle_webapp_data(update: Update, context):
    try:
        data = json.loads(update.message.web_app_data.data)
        rooms = data.get("rooms", [])
        totals = data.get("totals", {})

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Смета"

        headers = ["Помещение", "Фактура", "Вариант", "Работа ($)", "Материалы (расход, кг)", "Детали"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        for room in rooms:
            ws.append([
                room["name"],
                room["texture"],
                room.get("variant", ""),
                room["work"],
                room["mat"],
                room.get("details", "")
            ])

        ws.append([])
        ws.append(["ИТОГО", "", "", totals.get("work", "0"), totals.get("mat", "0"), ""])

        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_length + 4, 40)

        file = BytesIO()
        wb.save(file)
        file.seek(0)

        await update.message.reply_document(
            document=file,
            filename="smeta.xlsx",
            caption="📋 Ваша смета готова"
        )
    except Exception as e:
        await update.message.reply_text(f"Ошибка при создании сметы: {e}")

if __name__ == "__main__":
    app = Application.builder().token(TOKEN).build()
    app.add_handler(MessageHandler(filters.StatusUpdate.WEB_APP_DATA, handle_webapp_data))
    print("Бот запущен и ждёт данные из Mini App...")
    app.run_polling()
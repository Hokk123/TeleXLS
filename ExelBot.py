import telebot
from config import TOKEN
import pandas as pd
from datetime import datetime

bot = telebot.TeleBot(TOKEN)

# Получаем текущий год и месяц
year_now = datetime.now().year
month_now = datetime.now().month

# Загружаем данные из Excel-таблицы
def load_orders():
    # Укажите путь к вашей таблице Excel
    excel_file = 'test.xlsx'

    # Имя листа в таблице, где хранятся данные
    sheet_name = 'Лист1'

    # Загружаем данные из Excel-таблицы в DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df['номер'] = df['номер'].astype(str)
    return df

@bot.message_handler(commands=['start', 'help'])
def help(message: telebot.types.Message):
    text = 'Добро пожаловать! Этот бот предназначен для предоставления вам данных по реализации проекта системы связи и телекоммуникаций (ССиТ) ООО «Газпром межрегионгаз» (МРГ): номер записи'
    bot.reply_to(message, text)

# @bot.message_handler()
# def send_welcome(message):
#     bot.send_message(message.chat.id, f"Welcome, {message.chat.username}")

# Обработчик текстовых сообщений
@bot.message_handler(func=lambda message: True)
def handle_text(message):
    try:
        val = int(message.text)

        order_number = message.text.strip()

        # Загружаем данные из Excel-таблицы
        orders_df = load_orders()

        # Ищем данные по номеру
        status = orders_df.loc[orders_df['номер'] == order_number, 'данные'].values

        if len(status) > 0:
            bot.reply_to(message, f"{year_now} {month_now} данные {order_number}: {status[0]}")
        else:
            bot.reply_to(message, f"данные с номером {order_number} не найдены.")
            
    except ValueError:
        print("That's not an int!")

# Запускаем бота
bot.polling()


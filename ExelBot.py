import telebot
from config import keys, TOKEN
import pandas as pd

bot = telebot.TeleBot(TOKEN)


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
    text = 'Чтобы начать работу, введите команду боту в следующем формате:имя файла в формате xls'
    bot.reply_to(message, text)

# @bot.message_handler()
# def send_welcome(message):
#     bot.send_message(message.chat.id, f"Welcome, {message.chat.username}")

# Обработчик текстовых сообщений
@bot.message_handler(func=lambda message: True)
def handle_text(message):
    order_number = message.text.strip()

    # Загружаем данные из Excel-таблицы
    orders_df = load_orders()

    # Ищем статус заказа по номеру
    status = orders_df.loc[orders_df['номер'] == order_number, 'данные'].values

    if len(status) > 0:
        bot.reply_to(message, f"данные {order_number}: {status[0]}")
    else:
        bot.reply_to(message, f"данные с номером {order_number} не найдены.")

# Запускаем бота
bot.polling()


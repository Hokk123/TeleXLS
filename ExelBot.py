import telebot
from telebot import types
from config import TOKEN
import pandas as pd
from datetime import datetime

bot = telebot.TeleBot(TOKEN)

# Получаем текущий год и месяц
year_now = datetime.now().year
month_now = datetime.now().month

# Загружаем данные из Excel-таблицы
def load_data():
    # Укажите путь к вашей таблице Excel
    excel_file = 'mrg.xlsx'

    # Имя листа в таблице, где хранятся данные
    sheet_name = 'Лист1'

    # Загружаем данные из Excel-таблицы в DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    return df
    
# Стартовое приветствие
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton("Статистика проекта")
    btn2 = types.KeyboardButton("Планы проекта")
    markup.add(btn1, btn2)
    bot.send_message(message.chat.id, text="Добро пожаловать! Этот бот предназначен для предоставления вам данных по реализации проекта системы связи и телекоммуникаций (ССиТ) ООО «Газпром межрегионгаз» (МРГ):".format(message.from_user), reply_markup=markup)

# @bot.message_handler()
# def send_welcome(message):
#     bot.send_message(message.chat.id, f"Welcome, {message.chat.username}")

# Обработчик текстовых сообщений
@bot.message_handler(content_types=['text'])
def handle_text(message):

    # Загружаем данные из DataFrame
    data_df = load_data()

    if(message.text == "Статистика проекта"):
        # Выводим данные - Всего
        rows = len(data_df.index) #Количество строк в таблице
        total = data_df.iloc[(rows-1),2]
        if (total == 1):
            bot.send_message(message.chat.id, f"{year_now} {month_now} На текущий момент работы выполнены на {total} объекте")
        else:
            bot.send_message(message.chat.id, f"{year_now} {month_now} На текущий момент работы выполнены на {total} объектах")
    elif(message.text == "Планы проекта"):
        # Ищем данные по текущему месяцу
        
        bot.send_message(message.chat.id, text="Привеет.. Вывод планов в процессе создания!)")
        
        


# Запускаем бота
bot.polling()


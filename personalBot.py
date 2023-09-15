# -*- coding: utf-8 -*-
from email import message
import pytesseract
from PIL import Image
import openpyxl
from openpyxl import Workbook
# from pandas import test
from config import TOKEN
from BS import load_BS
from docxtpl import DocxTemplate
from thefuzz import fuzz
import sqlite3
from aiogram import Bot, Dispatcher, executor, types
from aiogram.dispatcher import FSMContext
from aiogram.types import ContentType
from aiogram.types import ReplyKeyboardRemove, \
    ReplyKeyboardMarkup, KeyboardButton, \
    InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.contrib.fsm_storage.memory import MemoryStorage
import random
import sys
import os
import datetime
from dateutil import parser

# описываем класс для добавления данных в базу
class DataSQL():
    def __init__(self) -> None:
        self.ID = None
        self.TBL = None
        self.DATE = None
        self.COUNT = None
        self.NOTE = None

class StateBot(StatesGroup):
    dateState = State()
    countState = State()
    noteState = State()

# создаем объект
data = DataSQL()

storage = MemoryStorage()

birthday = datetime.date(2000, 1, 1)

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'
file_path = 'D:\\python\\testPersonalBot\\'
doc = DocxTemplate('D:\\python\\testPersonalBot\\BS.docx')
bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=storage)
sql_file = "D:\\python\\testPersonalBot\\personal.db"

conn = sqlite3.connect(file_path + "personal.db") # или :memory: чтобы сохранить в RAM
cursor = conn.cursor()
conn.row_factory = sqlite3.Row

print("Начало работы бота.")

def godlet(god):
    '''Переводим цифру лет(int) в строку с нужным нам окончанием(str)'''
    return{
        god<=0:"лет",
        god%10==0:"лет",
        god%10==1:"год",
        god%10>1 and god%10<5:"года",
        god%10>4:"лет",
        god%100>10 and god%100<20:"лет"
        }[True]

def daymounth(day):
    '''Переводим цифру дней(int) в строку с нужным окончанием месяцев(str)'''
    return{
        day<=0:"месяцев",
        day==1:"месяц",
        day>1 and day<5:"месяца",
        day>4:"месяцев"
    }[True]

def get_maximum_rows(sheet_object):
    '''
    Функция нахождения максимального количества строк в эксель файле
    '''
    rows = 0
    for max_row, row in enumerate(sheet_object, 1):
        if not all(col.value is None for col in row):
            rows += 1
    return rows

# если мы добавляем новую таблицу, то сюда обязательно нужно внести предстваление с именем таблицы
list_nameButton = dict(Personal='Краткое резюме.', Otpusk = 'Отпуска.', ZaSvoySchet='Отпуск без содержания.', Bolnichniy='Больничные', Achievements='Ачивки', Projects='Проекты')

btnHello = KeyboardButton("/Personal")
btnRst = KeyboardButton('/rst')
btnTest = KeyboardButton('/add_to_DB')
btnINFO = KeyboardButton('/information')
greet_kb = ReplyKeyboardMarkup(resize_keyboard=True).row(btnHello, btnRst).add(btnTest, btnINFO)


#-------------------------------------------------------------------------
def get_keyboard_Personal(tp) -> types.InlineKeyboardMarkup:
    """
    Генерация клавиатуры из базы данных
    Выбор ID из базы Personal
    """
    cursor.execute("SELECT id, FIO FROM personal")
    pers_keyboard = cursor.fetchall()
    markup = types.InlineKeyboardMarkup()
    for id, Pers in pers_keyboard:
        markup.add(
            types.InlineKeyboardButton(
                Pers,
                callback_data=f'{tp}_'+str(id)),       # сдесь генерируется колбек дата в виде {tp}1-{tp}5
        )
    return markup


def get_keyboard_table(tp) -> types.InlineKeyboardMarkup:
    '''
    Берутся данные имена таблиц и данные из таблиц
    Имена таблиц 
    '''
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name != 'sqlite_sequence' AND name != 'Personal' AND name != 'Achievements' AND name!='Projects'")
    table_keyboard = cursor.fetchall()
    markup = types.InlineKeyboardMarkup()
    for id in table_keyboard:
        markup.add(
            types.InlineKeyboardButton(
                list_nameButton[str(id)[2:-3]],        # проверить соответствие имён таблиц в list_nameButton
                callback_data=f'{tp}_'+str(id)),       # сдесь генерируется колбек дата в виде tbl_Otpusk
        )
    return markup

#-------------------------------------------------------------------------
@dp.message_handler(commands=['start'])
async def process_hello(message: types.Message):
      await bot.send_message(message.from_user.id, 'Привет\nЧто будем делать?',
                             reply_markup=greet_kb)

@dp.message_handler(commands=['add_to_DB'])
async def process_help(message: types.Message):
    await bot.send_message(message.from_user.id, 'Для кого добавить данные?', reply_markup=get_keyboard_Personal('add'))

@dp.callback_query_handler(text='btn1')
async def firs_test(callback : types.CallbackQuery):
    test_text = ''
    cursor.execute("SELECT * FROM personal")
    text = cursor.fetchall()
    for i in text:
        for id in i:
            test_text = test_text + str(type(id)) + str(id) + '\n'
            print(test_text)
    # await callback.message.answer(test_text)
    await callback.message.edit_text(test_text, reply_markup=get_keyboard_Personal('test'))
    await callback.answer()

@dp.callback_query_handler(text='btn2')
async def firs_test(callback : types.CallbackQuery):
    test_text = ''
    cursor.execute("SELECT * FROM achievements")
    text = cursor.fetchall()
    for i in text:
        for id in i:
            test_text = test_text + str(type(id)) + str(id) + '\n'
            print(test_text)
    # await callback.message.answer(test_text)
    await callback.message.edit_text(test_text, reply_markup=get_keyboard_Personal('test1'))
    await callback.answer()


@dp.message_handler(commands='personal')
async def start_cmd_handler(message: types.Message):
    keyboard_markup = types.InlineKeyboardMarkup(row_width=3)
    # по умолчанию row_width равен 3, так что здесь мы можем его опустить
    # сохранено для ясности
    '''
    text_and_data = (
        ('Персонал!', 'personal'),
        ('Отпуска', 'otpuska'),
        ('ФИО', 'FIO')
    )
    # в реальной жизни для callback_data следует использовать фабрику данных обратного вызова
    # здесь для простоты используется необработанная строка
    row_btns = (types.InlineKeyboardButton(text, callback_data=data) for text, data in text_and_data)
    keyboard_markup.row(*row_btns)
    '''
    keyboard_markup.add(types.InlineKeyboardButton(list_nameButton['Personal'],callback_data='personal'))
    keyboard_markup.add(types.InlineKeyboardButton(list_nameButton['Otpusk'],callback_data='otpusk'))
    keyboard_markup.add(types.InlineKeyboardButton(list_nameButton['ZaSvoySchet'],callback_data='ZaSvoySchet'))
    keyboard_markup.add(types.InlineKeyboardButton(list_nameButton['Bolnichniy'],callback_data='Bolnichniy'))
    keyboard_markup.add(types.InlineKeyboardButton(list_nameButton['Achievements'],callback_data='Achievements'))
    #keyboard_markup.add(
        # url buttons have no callback data
    #    types.InlineKeyboardButton('Ссылка на гугл!', url='https://google.ru/'),
    #)

    await message.answer("Что вывести?", reply_markup=keyboard_markup)

@dp.callback_query_handler(text='personal')
async def answer_FIO(query: types.CallbackQuery):
    await query.message.edit_text('Персонал подразделения', reply_markup=get_keyboard_Personal('fio'))

@dp.callback_query_handler(text='otpusk')
async def answer_Otpuska(query: types.CallbackQuery):
    await query.message.edit_text('Отпуска всего подразделения.', reply_markup=get_keyboard_Personal('otp'))

@dp.callback_query_handler(text='ZaSvoySchet')
async def answer_BS(query: types.CallbackQuery):
    await query.message.edit_text('Отпуска за свой счет.', reply_markup=get_keyboard_Personal('bs'))

@dp.callback_query_handler(text='Bolnichniy')
async def answer_BS(query: types.CallbackQuery):
    await query.message.edit_text('Больничные', reply_markup=get_keyboard_Personal('boln'))

#-------------- Генерация колбеков по генерируемым кнопкам ---------------
@dp.callback_query_handler(text_startswith="fio_")
async def callbacks_fio(call: types.CallbackQuery):
    # Парсим строку и извлекаем действие, например `fio_1` -> `1`
    action = call.data.split("_")[1]
    cursor.execute(f"SELECT * FROM personal WHERE id = {action}")
    text = cursor.fetchall()
    temp1 = datetime.datetime.now() - parser.parse(text[0][2], dayfirst = True)
    let1 = str(temp1.days//365) + " " + godlet(temp1.days//365)
    temp2 = datetime.datetime.now() - parser.parse(text[0][3], dayfirst = True)
    let2 = str(temp2.days//365) + " " + godlet(temp2.days//365) + " " + str(temp2.days%365//30) + " " + daymounth(temp2.days%365//30)
    await call.message.edit_text(f"Краткое резюме:\n _ФИО_: {text[0][1]} \n _Дата рождения_: {text[0][2]} ({let1} ) \n _Дата трудоустройства_: {text[0][3]} ({let2})", parse_mode= "Markdown")
    await call.answer()
        # Если бы мы не меняли сообщение, то можно было бы просто удалить клавиатуру
        # вызовом await call.message.delete_reply_markup().
        # Но т.к. мы редактируем сообщение и не отправляем новую клавиатуру, 
        # то она будет удалена и так.
        # await call.message.edit_text(f"Итого: {user_value}")

#-------------------------------------------------------------------------
@dp.callback_query_handler(text_startswith="otp_")
async def callbacks_otp(call: types.CallbackQuery):
    # Парсим строку и извлекаем действие, например `otp_1` -> `1`
    action = call.data.split("_")[1]
    print(action)
    cursor.execute(f"SELECT id, FIO, date_otpusk, count_otpusk, note_otpusk FROM personal LEFT JOIN Otpusk ON personal.ID=Otpusk.ID_otpusk WHERE id = {action}")
    text = cursor.fetchall()
    print(text)
    await call.message.edit_text(f"{text[0][1]} \n_Отпуск с:_ {text[0][2]} \n_Количество дней:_ {text[0][3]} \n_Коментарий:_ {text[0][4]}", parse_mode= "Markdown")
    await call.answer()

#-------------------------------------------------------------------------
@dp.callback_query_handler(text_startswith="bs_")
async def callbacks_bs(call: types.CallbackQuery):
    # Парсим строку и извлекаем действие, например `bs_1` -> `1`
    action = call.data.split("_")[1]
    cursor.execute(f"SELECT id, FIO, date_BS, count_BS, note_BS FROM personal LEFT JOIN ZaSvoySchet ON personal.ID=ZaSvoySchet.ID_BS WHERE id = {action}")
    text = cursor.fetchall()
    print(text)
    temp_text = f"{text[0][1]}\n"
    for i in range(len(text)):
        temp_text += f"За свой счет: с {text[i-1][2]} числа, на {text[i-1][3]} дней/дня\n Примечание: {text[i-1][4]}"
    await call.message.edit_text(temp_text, parse_mode= "Markdown")
    await call.answer()

#-------------------------------------------------------------------------
@dp.callback_query_handler(text_startswith="boln_")
async def callbacks_boln(call: types.CallbackQuery):
    # Парсим строку и извлекаем действие, например `bs_1` -> `1`
    action = call.data.split("_")[1]
    cursor.execute(f"SELECT id, FIO, date_Bolnichniy, count_Bolnichniy, note_Bolnichniy FROM personal LEFT JOIN Bolnichniy ON personal.ID=Bolnichniy.ID_Bolnichniy WHERE id = {action}")
    text = cursor.fetchall()
    print(text)
    temp_text = f"{text[0][1]}\n"
    for i in range(len(text)):
        temp_text += f"Больничный: с {text[i-1][2]} числа, на {text[i-1][3]} дней/дня\n Примечание: {text[i-1][4]}"
    await call.message.edit_text(temp_text, parse_mode= "Markdown")
    await call.answer()

#-------------------------------------------------------------------------
@dp.callback_query_handler(text_startswith="add_")
async def callbacks_add(call: types.CallbackQuery):
    data.ID = call.data.split("_")[1]
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name != 'sqlite_sequence'") # выбор из таблиц
    text = cursor.fetchall()
    temp_text = ""
    for i in range(len(text)):
        temp_text += "/" + str(text[i])[2:-3] + "\n"
    await call.message.edit_text('Выберите нужную категорию для добавления', reply_markup=get_keyboard_table('tbl'))
    await call.answer()

# В таблицу нужно заносить следующие данные:
# Таблица, ID(выбирается по ФИО), дата, количество дней, примечание
#-------------------------------------------------------------------------
@dp.callback_query_handler(text_startswith="tbl_")
async def callbacks_tbl(call: types.CallbackQuery):
    data.TBL = call.data.split("_")[1][2:-3]
    print("LOG", data.ID, data.TBL)
    text = cursor.fetchall()
    temp_text = ""
    for i in range(len(text)):
        temp_text += "/" + str(text[i])[2:-3] + "\n"
    await call.message.edit_text('Далее следует ввести Дату события. 📅 Как пример: \n 23.12.2023')
    await StateBot.dateState.set()        # установили машину состояний в dateState
    await call.answer()
#-------------------------------------------------------------------------
@dp.message_handler(lambda message: not message.text.isdigit(), state=StateBot.countState)
async def check_age(message: types.Message):
    await message.reply('Введите количество дней цифрой!')

@dp.message_handler(state=StateBot.dateState)
async def load_data(message: types.Message, state: FSMContext) -> None:
    #async with state.proxy() as data:      # по идее это локальное хранилище данных
    data.DATE = message.text

    await message.answer('⏳А теперь введи количество дней.')
    await StateBot.next()

@dp.message_handler(state=StateBot.countState)
async def load_count(message: types.Message, state: FSMContext) -> None:
    #async with state.proxy() as data:      # по идее это локальное хранилище данных
    data.COUNT = message.text
    await message.reply('🗒 Введите примечание.')
    await StateBot.next()

@dp.message_handler(state=StateBot.noteState)
async def load_note(message: types.Message, state: FSMContext) -> None:
    #async with state.proxy() as data:      # по идее это локальное хранилище данных
    data.NOTE = message.text
    if data.TBL == 'ZaSvoySchet':
        load_BS(doc, sql_file, data.ID, data.DATE, data.COUNT)
    print(type(data.ID), type(data.TBL), type(data.DATE), type(data.COUNT), type(data.NOTE))
    print(data.ID, data.TBL, data.DATE, data.COUNT, data.NOTE)
    cursor.execute(f'INSERT INTO {data.TBL} VALUES (?,?,?,?)', (int(data.ID), str(data.DATE), int(data.COUNT), str(data.NOTE)))
    conn.commit()
    temp = list_nameButton[data.TBL]
    await message.reply(f'В таблицу "{temp}" были введены данные: \n 👤 Пользователь ID: {data.ID} \n 📅 Дата: {data.DATE} \n ⏳ Количество дней: {data.COUNT} \n 🗒 Примечание: {data.NOTE} ')
    await message.answer_document(open('BSfull.docx', 'rb'))
    await state.finish()


#-------------------------------------------------------------------------
# обработка фото
@dp.message_handler(content_types=['photo'])
async def photo(message):
    await message.photo[-1].download('test.jpg')
    max_prcnt = 0
    photo = Image.open('test.jpg')
    rus_string = pytesseract.image_to_string(photo, lang='rus')
    print(rus_string)
    print('Обработка фото')
    # print(rus_string)
    str_split = rus_string.lower().split('\n')
    try:
        wb = openpyxl.load_workbook(filename = 'vzb.xlsx', read_only=True)
    except:
        print('Проверте файл!!!')
        input('Нажмите ENTER.')
        sys.exit()
    ws = wb.active
    max_rows = get_maximum_rows(ws)
    cell_range = ws['b3':'d'+str(max_rows)]
    for i in range(int(max_rows)-2):
        if cell_range[i][0].value == None:
            continue
        result_percent = fuzz.ratio(str_split[0], cell_range[i][0].value.lower())
        if max_prcnt < result_percent:
            max_prcnt = result_percent
            str_result = cell_range[i][0].value.lower()
            str_result1 = cell_range[i][1].value.lower()        # тут иногда возникала ошибка в связи с тем что в ячейке хранилось обычное число
        # if result_percent > 70:
    print(str_split[0].lower(), '==>', str_result)
    print(max_prcnt)
    print(str_result1)
    result = 'Процент совпадения = ' + str(max_prcnt) + '\n' + 'Вопрос: ' + str(str_result) + '\n Ответ = ' + str(str_result1)
    await bot.send_message(message.from_user.id, result)
    # await bot.send_message(message.from_user.id, rus_string)

#-------------------------------------------------------------------------
@dp.message_handler(commands=['information'])
async def cmd_answer(message: types.Message):
    name = message.from_user.full_name
    username = message.from_user.username
    username = username and f"@{username}"
    id = message.from_user.id
    link = message.from_user.username
    link = link and f"https://t.me/{link}"
    await bot.send_message(message.from_user.id, f"👤 <b>Имя :</b> <b>{name}</b>\n🔑 <b>Имя пользователя :</b> <b>{username  if username  else None}</b>\n💳 <b>Телеграм ID :</b> <b>{id}</b> \n🔗 <b>Ссылка :</b> <b><a href='tg://user?id={id}'>{link if link else 'Ваша ссылка'}</a></b>", parse_mode="HTML")

@dp.message_handler(commands=['rst'])
async def restart(event):
      python = sys.executable
      os.execl(python, python, * sys.argv)


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)


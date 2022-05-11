from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
import sqlite3
import openpyxl
import datetime

bot = Bot(token='5357337704:AAFW4d8Yq3Gwp3-2J0oOMmyIzrwJSbUVu8c')
dp = Dispatcher(bot)

connection = sqlite3.connect('Users.db')
cur = connection.cursor()

wb = openpyxl.load_workbook('schedule.xlsx')
sheets = wb.sheetnames
sheet = wb[sheets[0]]
fixed_schedule = False

week_days = {'Понедельник': 0, 'Вторник': 1, 'Среда': 2, 'Четверг': 3, 'Пятница': 4, 'Суббота': 5, 'Воскресенье': 6}
week = {0: (6, 12), 1: (12, 18), 2: (18, 24), 3: (24, 30), 4: (30, 36), 5: (36, 42), 6: (42, 48)}
groups = {'1бд1': 'C', '1бд3': 'D', '1бу1': 'E', '1бу2': 'F', '1бу3': 'G', '1вб1': 'H', '1вб2': 'I', '1вб3': 'J',
          '1ис1': 'K', '1ис3': 'L', '1пк1': 'M', '1пк2': 'N', '1са1': 'O', '1са3': 'P'}


@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    await bot.send_message(chat_id=message.chat.id, text="Привет!\nВыбери свою группу!")
    if not cur.execute(f'''select chat_id From user_info
                        where chat_id = {message.chat.id} ''').fetchall():
        cur.execute("INSERT INTO user_info(chat_id, groups, admin)"
                    "VALUES(?, ?, ?)",
                    (message.chat.id, 'None', False))
        connection.commit()


@dp.message_handler(commands=['help'])
async def help(message: types.Message):
    await bot.send_message(chat_id=message.chat.id, text="Я бот, который будет скидыват тебе рассписание каждый день!")


@dp.message_handler(text=['Изменить расписание', 'Поменять расписание'])
async def edit_schedule(message: types.Message):
    global fixed_schedule
    if cur.execute(f'''select admin From user_info
                            where chat_id = {message.chat.id} ''').fetchall()[0][0]:

        await bot.send_message(chat_id=message.chat.id, text="Напишите мне новое расписание")
        fixed_schedule = True
    else:
        await bot.send_message(chat_id=message.chat.id, text="Вы не являетесь администратором!")


@dp.message_handler(text=['Завтра', 'Рассписание на завтра'])
async def schedule(message: types.Message):
    text = ''
    try:
        cort = week[datetime.datetime.today().weekday() + 1]
    except IndexError:
        cort = week[0]
    group = groups[cur.execute(f'''select groups From user_info
                                    where chat_id = {message.chat.id} ''').fetchall()[0][0]]
    for i in range(*cort):
        cell = sheet[f'{group}{str(i)}']
        if str(cell.value) == 'None':
            text += f'Пара №{abs(cort[0] - i) + 1} ----------- \n'
        else:
            text += f'Пара №{abs(cort[0] - i) + 1} {str(cell.value)}\n'
    await bot.send_message(chat_id=message.chat.id, text=text)


@dp.message_handler(text=['Расписание', 'Рассписание на сегодня'])
async def schedule(message: types.Message):
    text = ''
    cort = week[datetime.datetime.today().weekday()]
    group = groups[cur.execute(f'''select groups From user_info
                                    where chat_id = {message.chat.id} ''').fetchall()[0][0]]
    for i in range(*cort):
        cell = sheet[f'{group}{str(i)}']
        if str(cell.value) == 'None':
            text += f'Пара №{abs(cort[0] - i) + 1} ----------- \n'
        else:
            text += f'Пара №{abs(cort[0] - i) + 1} {str(cell.value)}\n'
    await bot.send_message(chat_id=message.chat.id, text=text)


@dp.message_handler(text=['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'])
async def schedule(message: types.Message):
    text = ''
    cort = week[week_days[message.text]]
    group = groups[cur.execute(f'''select groups From user_info
                                    where chat_id = {message.chat.id} ''').fetchall()[0][0]]
    for i in range(*cort):
        cell = sheet[f'{group}{str(i)}']
        if str(cell.value) == 'None':
            text += f'Пара №{abs(cort[0] - i) + 1} ----------- \n'
        else:
            text += f'Пара №{abs(cort[0] - i) + 1} {str(cell.value)}\n'
    await bot.send_message(chat_id=message.chat.id, text=text)


@dp.message_handler(content_types=['text'])
async def text(message: types.Message):
    global fixed_schedule
    if message.text not in groups:
        if cur.execute(f'''select groups From user_info
                                where chat_id = {message.chat.id} ''').fetchall() == [('None',)]:
            await bot.send_message(chat_id=message.chat.id, text='Такой группы не существует, попробуйте еще раз!')
        if fixed_schedule:
            for i in cur.execute(f'''select chat_id From user_info''').fetchall():
                await bot.send_message(chat_id=i[0], text=message.text)
                fixed_schedule = False
    else:
        cur.execute(f'''UPDATE user_info
                    SET groups = '{message.text}'
                    WHERE chat_id = {message.chat.id} ''')
        connection.commit()
        await bot.send_message(chat_id=message.chat.id,
                               text=f'В данной беседе была установленна группа "{message.text}"')


if __name__ == '__main__':
    executor.start_polling(dp)

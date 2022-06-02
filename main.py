from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
import sqlite3
import openpyxl
import datetime
import aioschedule
import asyncio

bot = Bot(token='5357337704:AAFKs87RQe0U-6pmOaZlsC7qBYyM156NpC8')
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
          '1ис1': 'K', '1ис3': 'L', '1пк1': 'M', '1пк2': 'N', '1са1': 'O', '1са3': 'P', '2бд1': 'Q', '2бд2': 'R',
          '2бд3': 'S', '2бд4': 'T', '2бу1': 'U', '2бу2': 'V', '2бу3': 'W', '2вб1': 'X', '2вб3': 'Y', '2ис1': 'Z',
          '2ис2': 'AA', '2ис3': 'AB', '2пк1': 'AC', '2пк2': 'AD', '2са1': 'AE', '2са3': 'AF', '3бу1': 'AG',
          '3бу2': 'AH', '3ис1': 'AI', '3ис2': 'AJ', '3ис3': 'AK', '3ис4': 'AL', '3пк1': 'AM', '3пк2': 'AN',
          '3са1': 'AO', '3са2': 'AP', '3са3': 'AQ'}


@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    text = 'Привет! 👋' \
           '\nЯ бот, который будет отправлять тебе рассписание каждый день в 06:00 по МСК' \
           '\n1⃣Чтобы указать группу, в который ты учишься, тебе нужно просто отправить сообщение: <Название группы>' \
           '\n❗Например: 1вб2' \
           '\n2⃣Чтобы изменить группу, просто введи её название' \
           '\n❗Например: 1вб2' \
           '\n3⃣Если тебе нужно рассписание на конкретный день недели, то просто напиши об это мне: <День недели> ' \
           '\n❗Например: Четверг' \
           '\n4⃣У меня есть еще одна функция, если ты напишешь: <Завтра> или <Сегодня>, то я тебе пришлю рассписание на этот день' \
           '\n5⃣Tакже, если ты админ, то ты можешь кидать изменённое рассписание'
    await bot.send_message(chat_id=message.chat.id, text=text)
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
    text = 'Расписание для группы "' + str(cur.execute(f'''select groups From user_info
                                        where chat_id = {message.chat.id} ''').fetchall()[0][0]) + '"\n\n'
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


@dp.message_handler(text=['Расписание', 'Рассписание на сегодня', 'Сегодня'])
async def schedule(message: types.Message):
    text = 'Расписание для группы "' + str(cur.execute(f'''select groups From user_info
                                        where chat_id = {message.chat.id} ''').fetchall()[0][0]) + '"\n\n'
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
    cort = week[week_days[message.text]]
    text = 'Расписание для группы "' + str(cur.execute(f'''select groups From user_info
                                    where chat_id = {message.chat.id} ''').fetchall()[0][0]) + '"\n\n'
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


async def morning():  # функция рассылки кажое утро
    chat_id = cur.execute(
        f'''select chat_id From user_info''').fetchall()  # получаем всех юзеров из БД с вкл уведмолениями
    for i in chat_id:  # пробегаемся по всем найденым юзерам
        text = 'Расписание для группы "' + str(cur.execute(f'''select groups From user_info
                                                    where chat_id = {i} ''').fetchall()[0][0]) + '"\n\n'
        cort = week[datetime.datetime.today().weekday()]
        group = groups[cur.execute(f'''select groups From user_info
                                                where chat_id = {i} ''').fetchall()[0][0]]
        for i in range(*cort):
            cell = sheet[f'{group}{str(i)}']
            if str(cell.value) == 'None':
                text += f'Пара №{abs(cort[0] - i) + 1} ----------- \n'
            else:
                text += f'Пара №{abs(cort[0] - i) + 1} {str(cell.value)}\n'
        await bot.send_message(i, text)  # присылаем погоду


async def scheduler():  # функция отслеживания времени
    try:
        aioschedule.every().day.at("08:00").do(morning)  # проверка времени и срабаотываение функции
        while True:
            await aioschedule.run_pending()
            await asyncio.sleep(1)
    except Exception:
        pass


async def on_startup(dp):
    asyncio.create_task(scheduler())


if __name__ == '__main__':
    executor.start_polling(dp, on_startup=on_startup)

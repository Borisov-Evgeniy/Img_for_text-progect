import os
import pytesseract
from PIL import Image
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import InputFile
from aiogram.utils import executor
from docxtpl import DocxTemplate
from openpyxl import Workbook

# Путь к директории tessdata, где содержатся языковые пакеты
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'
# Путь к директории с исполняемым файлом Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

API_TOKEN = 'YOUR_TOKEN'
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)


class TextExtraction(StatesGroup):
    waiting_for_screenshot = State()
    waiting_for_output_type = State()


@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    await message.reply(f'Привет! Пришли мне изображение и я считаю с него текст.')


@dp.message_handler(content_types=[types.ContentType.PHOTO], state='*')
async def handle_screenshot(message: types.Message, state: FSMContext):
    photo = message.photo[-1]
    file_id = photo.file_id
    file = await bot.get_file(file_id)
    downloaded_file = await bot.download_file(file.file_path)
    image = Image.open(downloaded_file)

    text = pytesseract.image_to_string(image, lang='rus+eng')
    await state.update_data(text=text)
    await TextExtraction.waiting_for_output_type.set()

    keyboard = types.InlineKeyboardMarkup()
    button1 = types.InlineKeyboardButton(text="Word", callback_data="word")
    button2 = types.InlineKeyboardButton(text="TXT", callback_data="txt")
    button3 = types.InlineKeyboardButton(text="Excel", callback_data="excel")
    button4 = types.InlineKeyboardButton(text="Telegram", callback_data="telegram")

    keyboard.row(button1)
    keyboard.row(button2)
    keyboard.row(button3)
    keyboard.row(button4)

    await message.reply("Выберите формат вывода:", reply_markup=keyboard)


@dp.callback_query_handler(lambda c: c.data in ['word', 'txt', 'excel', 'telegram'],
                           state=TextExtraction.waiting_for_output_type)
async def handle_output_type(callback_query: types.CallbackQuery, state: FSMContext):
    output_type = callback_query.data
    data = await state.get_data()
    text = data.get('text', '')

    if output_type == 'word':
        filename = 'output.docx'
        doc = DocxTemplate("templates/base.docx")  # путь к базовому .docx шаблону
        context = {'text': text}
        doc.render(context)
        doc.save(filename)
        await bot.send_document(callback_query.from_user.id, InputFile(filename))
    elif output_type == 'txt':
        filename = 'output.txt'
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(text)
        await bot.send_document(callback_query.from_user.id, InputFile(filename))
    elif output_type == 'excel':
        filename = 'output.xlsx'
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.cell(1, 1, value=text)
        workbook.save(filename)
        await bot.send_document(callback_query.from_user.id, InputFile(filename))
    elif output_type == 'telegram':
        await bot.send_message(callback_query.from_user.id, text)

    await state.finish()
    await bot.answer_callback_query(callback_query.id)


if __name__ == '__main__':
    dp.register_message_handler(send_welcome, commands=['start'])
    executor.start_polling(dp, skip_updates=True)
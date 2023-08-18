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


# Указываем путь к директории tessdata, где содержатся языковые пакеты
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

# Указываем путь к директории с исполняемым файлом Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

API_TOKEN = '6100027035:AAEi16gmd2eLPKvOJ9XSM_ebIq65dcznI9g'
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)


class TextExtraction(StatesGroup):
    waiting_for_screenshot = State()
    waiting_for_output_type = State()


@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    await message.reply(f'Hello! Send me a screenshot of text to extract the text from it.')


@dp.message_handler(content_types=[types.ContentType.PHOTO], state='*')
async def handle_screenshot(message: types.Message, state: FSMContext):
    photo = message.photo[-1]
    file_id = photo.file_id
    file = await bot.get_file(file_id)
    downloaded_file = await bot.download_file(file.file_path)
    image = Image.open(downloaded_file)

    # Используем параметр lang для установки языка русский
    text = pytesseract.image_to_string(image, lang='rus+eng')

    await state.update_data(text=text)
    await TextExtraction.waiting_for_output_type.set()
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.add(types.KeyboardButton('Word'), types.KeyboardButton('TXT'))
    keyboard.add(types.KeyboardButton('Excel'), types.KeyboardButton('Telegram'))
    await message.reply("Choose the output format:", reply_markup=keyboard)


@dp.message_handler(lambda message: message.text in ['Word', 'TXT', 'Excel', 'Telegram'],
                    state=TextExtraction.waiting_for_output_type)
async def handle_output_type(message: types.Message, state: FSMContext):
    output_type = message.text
    data = await state.get_data()
    text = data.get('text', '')
    if output_type == 'Word':
        filename = 'output.docx'
        doc = DocxTemplate("templates/base.docx")  # path to the base .docx template
        context = {'text': text}
        doc.render(context)
        doc.save(filename)
        await bot.send_document(message.chat.id, InputFile(filename))
    elif output_type == 'TXT':
        filename = 'output.txt'
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(text)
        await bot.send_document(message.chat.id, InputFile(filename))
    elif output_type == 'Excel':
        filename = 'output.xlsx'
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.cell(1, 1, value=text)
        workbook.save(filename)
        await bot.send_document(message.chat.id, InputFile(filename))
    elif output_type == 'Telegram':
        await message.reply(text)
    await state.finish()


if __name__ == '__main__':
    dp.register_message_handler(send_welcome, commands=['start'])
    executor.start_polling(dp, skip_updates=True)

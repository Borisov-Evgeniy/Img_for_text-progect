Этот бот разработан для извлечения текста из изображений и документов, а также конвертации этого текста в разные форматы. Он использует библиотеки Aiogram, pytesseract, PIL, docxtpl, openpyxl и работает с Telegram API.

## Требования

- Python 3.8 или выше
- Установленные библиотеки, перечисленные в начале скрипта (pytesseract, aiogram, PIL, docxtpl, openpyxl)
- Tesseract OCR

## Установка Tesseract OCR

Для успешной работы бота, вам также необходимо установить и настроить Tesseract OCR - программу для распознавания текста на изображениях.

### Windows:

1. Скачайте установщик Tesseract для Windows по [ссылке](https://github.com/tesseract-ocr/tesseract/releases).
2. Запустите установщик и следуйте инструкциям на экране.
3. После установки, добавьте путь к директории с исполняемым файлом Tesseract OCR в переменную среды `PATH`.

### Linux (Ubuntu):

1. Откройте терминал и выполните команду `sudo apt-get install tesseract-ocr`.

## Как использовать

1. Зарегистрируйте своего бота на [Telegram BotFather](https://core.telegram.org/bots#botfather) и получите API токен.
2. Установите все необходимые библиотеки с помощью `pip install -r requirements.txt`.
3. Замените значение переменной `API_TOKEN` на ваш полученный API токен.
4. Замените пути к директориям Tesseract OCR и исполняемому файлу Tesseract OCR на свои в переменных `os.environ['TESSDATA_PREFIX']` и `pytesseract.pytesseract.tesseract_cmd` соответственно.
5. Создайте необходимые папки и файлы для сохранения документов (Word шаблон, например).
6. Запустите программу: `python your_script_name.py`.

## Примечание

- Проверьте совместимость версий библиотек. Возможно, потребуется обновление или замена некоторых компонентов.
- Этот бот ориентирован на использование локальных путей и может потребовать доработки для размещения на сервере.

## Благодарности

Этот бот создан благодаря использованию открытых исходных кодов следующих библиотек:
- [Aiogram](https://github.com/aiogram/aiogram)
- [pytesseract](https://github.com/madmaze/pytesseract)
- [PIL](https://github.com/python-pillow/Pillow)
- [docxtpl](https://github.com/elapouya/python-docx-template)
- [openpyxl](https://github.com/openpyxl/openpyxl)
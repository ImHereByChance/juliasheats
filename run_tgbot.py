import os
import urllib
from main import convert_file

from aiogram import Bot, Dispatcher, executor, types


TOKEN = os.getenv('TG_TOKEN')


bot = Bot(token=TOKEN)
dp = Dispatcher(bot)


@dp.message_handler(content_types=['document'])
async def scan_message(message: types.Message):
    document_id = message.document.file_id
    file_info = await bot.get_file(document_id)
    url_path_to_file = file_info.file_path
    source_file = message.document.file_name
    urllib.request.urlretrieve(
        f'https://api.telegram.org/file/bot{TOKEN}/{url_path_to_file}',
        f'./{source_file}')
    try:
        final_file = convert_file(source_file)
    except Exception as e:
        await message.reply(f"Sorry, I couldn't convert that.\n {e}")
        raise(e)

    with open(f'{final_file}', 'rb') as file:
        await message.reply_document(file)
        
# @dp.message_handler(content_types=['document'])
# async def scan_message(message: types.Message):
#     document_id = message.document.file_id
#     file_info = await bot.get_file(document_id)
#     url_path_to_file = file_info.file_path
#     source_file = message.document.file_name
#     urllib.request.urlretrieve(f'https://api.telegram.org/file/bot{TOKEN}/{url_path_to_file}', f'./{source_file}')
    
#     final_file = convert_file(source_file)
#     with open(f'{final_file}', 'rb') as file:
#         await message.reply_document(file)


def main():
    executor.start_polling(dp, skip_updates=True)


if __name__ == '__main__':
    main()

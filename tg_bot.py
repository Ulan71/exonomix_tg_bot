import asyncio


from aiogram import Bot, Dispatcher, types, F
from aiogram.filters.command import Command
from aiogram.methods import DeleteWebhook
from aiogram.types import CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.filters.command import CommandObject
from aiogram.filters.state import StatesGroup, State

TOKEN = "8444253852:AAGGbPsm-Ojpu3KrGd9fxKtHH2CNKFJU5Qk"

bot = Bot(token=TOKEN)

dp = Dispatcher()


class UserState(StatesGroup):
    lobby = State()
    waiting_for_file = State()


@dp.message(Command("start"))
async def start(message: types.Message):
    await message.answer("Yes")


@dp.message(F.document)
async def get_user_file(message: types.Message):
    await bot.download_file()


if __name__ == "__main__":
    asyncio.run(dp.start_polling(bot))

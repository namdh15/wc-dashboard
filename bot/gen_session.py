"""
Chạy 1 lần tại máy để lấy SESSION_STRING.
Sau đó paste giá trị đó vào biến môi trường trên Railway/Render.
"""
import asyncio
from telethon import TelegramClient
from telethon.sessions import StringSession

API_ID   = 34813147
API_HASH = "ec672e74a5ffa437c7eea4a144c0b423"
PHONE    = "+84 869 322 628"

async def main():
    async with TelegramClient(StringSession(), API_ID, API_HASH) as client:
        await client.start(phone=PHONE)
        print("\n=== SESSION STRING (copy toàn bộ dòng dưới) ===")
        print(client.session.save())
        print("===============================================\n")

asyncio.run(main())

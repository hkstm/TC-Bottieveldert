import asyncio
import discord
import argparse
import MessageManager
from pathlib import Path
credentials_path = Path('Credentials.py')
if credentials_path.is_file():
    from Credentials import TOKEN, GENERAL_CHANNEL, BOTTEST_CHANNEL

parser = argparse.ArgumentParser(description='Process some integers.')
parser.add_argument('-t', '--token', type=str, nargs='?', default='',
                    help='The token for your Discord bot, you could check https://discordpy.readthedocs.io/en/latest/discord.html for more information')
parser.add_argument('-c', '--channel', type=int, nargs='?', default=0,
                    help='The channel you want your bot to send messages to, you could check https://discordpy.readthedocs.io/en/latest/discord.html for more information')
args = parser.parse_args()

TOKEN = TOKEN if args.token == '' else args.token
CHANNEL = BOTTEST_CHANNEL if args.channel == 0 else args.channel

client = discord.Client()


async def time_check():
    await client.wait_until_ready()
    while not client.is_closed():
        msg = MessageManager.get_message()
        if len(msg) > 0:  # message length 0 indicates either buggy message or more likely that it's not time to send a message
            channel = client.get_channel(CHANNEL)
            try:
                await channel.send(msg)
            except discord.errors.HTTPException:
                print(f'HTTPException, message: {msg}')
            time = (5 * 60)  # check every minute if a message should be send, this to allow time offset in get_message to function
        else:
            time = (1 * 60)  # check every minute if a message should be send, this to allow time offset in get_message to function
        await asyncio.sleep(time)


async def update_data():
    await client.wait_until_ready()
    while not client.is_closed():
        MessageManager.update_table_csv()
        MessageManager.update_fixtures_csv()
        time = (1 * 60 * 60 * 24)  # update data daily
        await asyncio.sleep(time)


client.loop.create_task(time_check())
client.loop.create_task(update_data())
client.run(TOKEN)

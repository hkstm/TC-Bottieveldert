import asyncio
import discord
import argparse
import MessageManager
from Credentials import TOKEN, GENERAL_CHANNEL, BOTTEST_CHANNEL

parser = argparse.ArgumentParser(description='Process some integers.')
parser.add_argument('-t', '--token', type=str, nargs='?', default='',
                    help='The token for your Discord bot, you could check https://discordpy.readthedocs.io/en/latest/discord.html for more information')
args = parser.parse_args()

TOKEN = TOKEN if args.token == '' else args.token
client = discord.Client()


async def time_check():
    await client.wait_until_ready()
    while not client.is_closed():
        msg = MessageManager.get_message()
        if len(msg) > 0:  # message length 0 indicates either buggy message or more likely that it's not time to send a message
            channel = client.get_channel(GENERAL_CHANNEL)
            try:
                await channel.send(msg)
            except discord.errors.HTTPException:
                print(f'HTTPException, message: {msg}')
        time = 1 * 60  # check every minute if a message should be send, this to allow time offset in get_message to function
        await asyncio.sleep(time)


client.loop.create_task(time_check())
client.run(TOKEN)

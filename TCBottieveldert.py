import asyncio
import discord
import argparse
import MessageManager
from Credentials import TOKEN, CHANNEL

parser = argparse.ArgumentParser(description='Process some integers.')
parser.add_argument('-t', '--token', type=str, nargs='?', default='',
                    help='The token for your Discord bot, you could check https://discordpy.readthedocs.io/en/latest/discord.html for more information')
args = parser.parse_args()

TOKEN = TOKEN if args.token == '' else args.token
client = discord.Client()


async def time_check():
    await client.wait_until_ready()
    general_channel = CHANNEL
    while not client.is_closed():
        if MessageManager.is_message_event_time():
            channel = client.get_channel(general_channel)
            await channel.send(MessageManager.get_message())
        time = 1 * 60  # check every minute if a message should be send, this to allow time offset in get_message to function
        await asyncio.sleep(time)


client.loop.create_task(time_check())
client.run(TOKEN)

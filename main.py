import os
import time
import asyncio
import json
import math
import win32com.client

from datetime import datetime

from core.database.plugin import PluginBaseModel
from core.database.group import GroupBaseModel
from core.database.messages import *
from core import send_to_console_channel, Message, Chain, AmiyaBotPluginInstance, bot as main_bot, OneBot11Instance

from amiyabot.database import *
from core.database.bot import Admin

curr_dir = os.path.dirname(__file__)
scheduler = win32com.client.Dispatch('Schedule.Service')


class TimerPluginInstance(AmiyaBotPluginInstance):
    ...


bot = TimerPluginInstance(
    name='amiyabot-schedule-controller',
    version='1.0',
    plugin_id='amiyabot-schedule-controller',
    plugin_type='',
    description='控制Windows计划任务程序',
    document=f'{curr_dir}/README.md',
    global_config_schema=f'{curr_dir}/config_schema.json',
    global_config_default=f'{curr_dir}/config_default.yaml'
)


def get_task(task_name):
    scheduler.Connect()
    root_folder = scheduler.GetFolder('\\')
    task = root_folder.GetTask(task_name)
    return task


@bot.on_message(group_id='remind', keywords=['启用任务'], level=5, direct_only=True)
async def _(data: Message):
    if not bool(Admin.get_or_none(account=data.user_id)):
        return Chain(data).text('抱歉，该功能只能由兔兔管理员操作~')

    try:
        _, name = data.text.split(' ', 1)
        get_task(name).Enabled = True
        return Chain(data).text(f'成功启用任务{name}')
    except Exception as e:
        return Chain(data).text(f'操作失败: {e}')


@bot.on_message(group_id='remind', keywords=['禁用任务'], level=5, direct_only=True)
async def _(data: Message):
    if not bool(Admin.get_or_none(account=data.user_id)):
        return Chain(data).text('抱歉，该功能只能由兔兔管理员操作~')

    try:
        _, name = data.text.split(' ', 1)
        get_task(name).Enabled = False
        return Chain(data).text(f'成功禁用任务{name}')
    except Exception as e:
        return Chain(data).text(f'操作失败: {e}')


@bot.on_message(group_id='remind', keywords=['运行任务'], level=5, direct_only=True)
async def _(data: Message):
    if not bool(Admin.get_or_none(account=data.user_id)):
        return Chain(data).text('抱歉，该功能只能由兔兔管理员操作~')

    try:
        _, name = data.text.split(' ', 1)
        get_task(name).Run(0)
        return Chain(data).text(f'成功运行任务{name}')
    except Exception as e:
        return Chain(data).text(f'操作失败: {e}')

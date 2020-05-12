import requests
import jsonpickle
import random
import pytz
from datetime import datetime, timedelta
from openpyxl import load_workbook
from Insults import insult_firstword, insult_secondword, insult_thirdword

eltp_tz = pytz.timezone('Europe/London')  # Timezone that game times are given in are in Lodon
ownteam = 'TC Ballieveldert'
date_format = '%d/%m'
time_format = '%A:%H:%M'

matchday_msg_timeoffset_mins_list = [8 * 60, (1 * 60) + 15, 5]  # timeoffset in minutes
scrim_msg_timeoffset_mins_list = [8 * 60, 5]  # timeoffset in minutes

scrim_offset = -5


def is_dst(tz):
    now = pytz.utc.localize(datetime.utcnow())
    return now.astimezone(tz).dst() != timedelta(0)


def conv_eventinfo_to_dict(datetime, type):
    return {
        'datetime': datetime,
        'type': type,
    }


def make_event_list():
    global event_datetime_list

    ws_dataonly = load_workbook(filename='data/minorsfixtures.xlsx', data_only=True)['Minors Fixtures']

    event_datetime_list = []
    event_datetime_list.extend([conv_eventinfo_to_dict(datetime(year=2020, month=5, day=12,
                                                               hour=(19 if is_dst(eltp_tz) else 20), minute=0,
                                                               tzinfo=pytz.utc), 'match')])  # fo)
    for row_idx, row in enumerate(ws_dataonly.rows):  # go through all the rows of the fixtures tab
        if isinstance(row[0].value,
                      datetime):  # find the row that matches the current date, e.g. day on which map is played
            match_datetime = datetime(year=row[0].value.year, month=row[0].value.month, day=row[0].value.day,
                                      hour=(19 if is_dst(eltp_tz) else 20), minute=0,
                                      tzinfo=pytz.utc)  # for calculations it is recommended to store times in utc, so we have to correct for that wrt actual game times
            scrim_datetime = match_datetime + timedelta(days=scrim_offset)  # scrims 5 days before actual game e.g. Wednesdays
            event_datetime_list.extend([conv_eventinfo_to_dict(match_datetime, 'match'),
                                        conv_eventinfo_to_dict(scrim_datetime, 'scrim')])
    with open('data/eventinfo.json', 'wb') as outfile:
        outfile.write(jsonpickle.encode(event_datetime_list).encode("utf-8"))


make_event_list()

# event_datetime_list = None
with open('data/eventinfo.json',
          'rb') as infile:  # if the general sheet fixture is updated/made for the first time run update_fixtures_csv
    event_datetime_list = jsonpickle.decode(infile.read().decode("utf-8"))


def update_fixtures_csv():
    minorsfixtures_url = 'https://docs.google.com/spreadsheets/d/1ruIpfqYwHyH17tvN6PVEYqYJm8geUoqgdsGgzm-BjhM/export?format=xlsx&id=1ruIpfqYwHyH17tvN6PVEYqYJm8geUoqgdsGgzm-BjhM&gid=92829842'
    open('data/minorsfixtures.xlsx', 'wb').write(requests.get(url=minorsfixtures_url).content)
    make_event_list()


def update_table_csv():
    minorstable_url = 'https://docs.google.com/spreadsheets/d/1ruIpfqYwHyH17tvN6PVEYqYJm8geUoqgdsGgzm-BjhM/export?format=xlsx&id=1ruIpfqYwHyH17tvN6PVEYqYJm8geUoqgdsGgzm-BjhM&gid=601872256'
    open('data/minorstable.xlsx', 'wb').write(requests.get(url=minorstable_url).content)


def get_ranking(team):
    ws_table_dataonly = load_workbook(filename='data/minorstable.xlsx', data_only=True)['Minors Table']
    for row in ws_table_dataonly.iter_rows():
        if team == row[2].value:
            return row[0].value


def ranking_to_placement(ranking):
    if ranking == 1:
        return str(ranking) + 'st'
    elif ranking == 2:
        return str(ranking) + 'nd'
    else:
        return str(ranking) + 'th'


def create_msg(a_event, important_offset_mins, timeoffset_mins):
    current_date = datetime.strftime(
        (a_event['datetime'] if a_event['type'] == 'match' else a_event['datetime'] + timedelta(days=-scrim_offset)),
        date_format)
    ws_fixt = load_workbook(filename='data/minorsfixtures.xlsx', data_only=False)['Minors Fixtures']
    ws_fixt_dataonly = load_workbook(filename='data/minorsfixtures.xlsx', data_only=True)['Minors Fixtures']

    daterowinfo_list = []
    opponent = ''
    teamcol1 = 1  # first column with team name on https://docs.google.com/spreadsheets/d/1ruIpfqYwHyH17tvN6PVEYqYJm8geUoqgdsGgzm-BjhM/edit#gid=92829842
    teamcol2 = 7  # second column with team name on https://docs.google.com/spreadsheets/d/1ruIpfqYwHyH17tvN6PVEYqYJm8geUoqgdsGgzm-BjhM/edit#gid=92829842
    for row_idx, row in enumerate(ws_fixt_dataonly.rows):  # go through all the rows of the fixtures tab
        if isinstance(row[0].value, datetime) and datetime.strftime(row[0].value,
                                                                    '%d/%m') == current_date:  # find the row that matches the current date, e.g. day on which map is played

            for col_idx, curr_cell in enumerate(row):
                if curr_cell.value is not None:  # add all cells that are not None to daterowinfo_list
                    daterowinfo_list.append(
                        (curr_cell.value, ws_fixt.cell(row=row_idx + 1,
                                                       column=col_idx + 1).value))  # adding cell value from curr_cell.value and formula from ws_fixt.cell to get link to map
            for row_opp_idx in range(row_idx + 2,
                                     row_idx + 6):  # from the row in which the date of the week and maps are specified, the actual matches are gives in this range of rows
                if ws_fixt.cell(row=row_opp_idx + 1,
                                column=teamcol1).value == ownteam:  # if we are team 1 then our opponent for this week is team 2
                    opponent = ws_fixt.cell(row=row_opp_idx + 1, column=teamcol2).value
                    break
                elif ws_fixt.cell(row=row_opp_idx + 1, column=teamcol2).value == ownteam:  # and vice versa
                    opponent = ws_fixt.cell(row=row_opp_idx + 1, column=teamcol1).value
                    break
            break

    own_ranking = ranking_to_placement(get_ranking(ownteam))
    opponent_ranking = ranking_to_placement(get_ranking(opponent))
    week, _ = daterowinfo_list[2]
    map1_name, map1_link = daterowinfo_list[1]
    map1_link = map1_link.split('"')[1] if len(map1_link.split('"')) > 1 else map1_link
    map2_name, map2_link = daterowinfo_list[3]
    map2_link = map2_link.split('"')[1] if len(map2_link.split('"')) > 1 else map2_link
    map3_name, map3_link = daterowinfo_list[4]
    map3_link = map3_link.split('"')[1] if len(map3_link.split('"')) > 1 else map3_link

    hours_val = timeoffset_mins // 60
    mins_val = timeoffset_mins % 60

    msg = ''
    if a_event['type'] == 'match':
        if timeoffset_mins == important_offset_mins:
            msg = f"@everyone MATCHDAY - Today we're playing {opponent} (ranked {opponent_ranking}) on {map1_name}, {map2_name} & {map3_name} at 8pm BST" \
                  f" that's in less than {hours_val} hours and {mins_val} minutes." \
                  f" Lets beat those {random.choice(insult_firstword)} {random.choice(insult_secondword)} {random.choice(insult_thirdword)}!!" \
                  f" Please be on at 7pm BST for scrims." \
                  f"\nWe're currently in {week.lower()} out of 7 weeks of the regular season and are ranked {own_ranking}."
        else:
            msg = f" MATCHDAY REMINDER - Today we're playing {opponent} (ranked {opponent_ranking}) on {map1_name}, {map2_name} & {map3_name} at 8pm BST" \
                  f" that's in less than {hours_val} hours and {mins_val} minutes." \
                  f" Please be on at 7pm BST for scrims."

    elif a_event['type'] == 'scrim':
        msg = f"PRACTICE - Playing {opponent} on {map1_name}, {map2_name} & {map3_name} this week" \
              f" Please be on at 8pm BST for practice." \
              f" that's in less than {hours_val} hours and {mins_val} minutes."
    if len(msg) > 0 and (a_event['type'] == 'scrim' or timeoffset_mins == important_offset_mins):
        msg += f"\n{map1_name}: {map1_link}," \
               f" {map2_name}: {map2_link}," \
               f" {map3_name}: {map3_link}"
    else:
        print(f'message empty, event: {a_event}')
    return msg


def get_message():
    current_time = datetime.strftime(datetime.utcnow(), time_format)
    current_date = datetime.strftime(datetime.utcnow(), date_format)

    current_time_mins = int(current_time.split(':')[1]) * 60 + int(
        current_time.split(':')[2])  # time_format = '%A:%H:%M'
    for event in event_datetime_list:
        event_time = datetime.strftime(event['datetime'], time_format)
        event_date = datetime.strftime(event['datetime'], date_format)

        if event_date == current_date:
            if event['type'] == 'match':
                for timeoffset_mins in matchday_msg_timeoffset_mins_list:
                    event_msg_time_mins = int(event_time.split(':')[1]) * 60 + int(
                        event_time.split(':')[2]) - timeoffset_mins
                    if current_time_mins == event_msg_time_mins:
                        return create_msg(event, matchday_msg_timeoffset_mins_list[0], timeoffset_mins)
            elif event['type'] == 'scrim':
                for timeoffset_mins in scrim_msg_timeoffset_mins_list:
                    event_msg_time_mins = int(event_time.split(':')[1]) * 60 + int(
                        event_time.split(':')[2]) - timeoffset_mins
                    if current_time_mins == event_msg_time_mins:
                        return create_msg(event, -1, timeoffset_mins)
    return ''

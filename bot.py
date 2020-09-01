# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib oauth2client

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# pip install python-telegram-bot
# pip install schedule

import telegram
from telegram.ext import (Updater)
import logging
import schedule
import time
import datetime
from params import bottoken, channel, group, SPREADSHEET_ID, RANGE_NAME

bot = telegram.Bot(token=bottoken)

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.DEBUG)
logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']


def checker():
    message = task()
    if message:
        bot.send_message(chat_id=group, text='<i>This announcement will be sent at 5pm today</i>',
                         parse_mode=telegram.ParseMode.HTML)
        bot.send_message(chat_id=group, text=message,
                         parse_mode=telegram.ParseMode.HTML, disable_web_page_preview=True)


def sender():
    message = task()
    if message:
        bot.send_message(chat_id=channel, text=message,
                         parse_mode=telegram.ParseMode.HTML, disable_web_page_preview=True)


def task():
    values = getsheet()
    for row in values:
        try:
            date = row[0]
        except:
            continue
        if datetime.datetime.now().strftime("%d/%m/%y") == date:
            try:
                intro = row[1].strip()
            except:
                continue
            try:
                day = row[2].strip()
                title = row[3].strip()
                if not day:
                    raise ValueError
                if not title:
                    raise ValueError
                try:
                    speaker = row[4].strip()
                    if not speaker:
                        raise ValueError
                    speaker = ' by <b>{}</b>'.format(speaker)
                except:
                    speaker = ''
                part1 = ' <b>{}</b>, we will be having <b>{}</b>{}.'.format(
                    day, title, speaker)
            except:
                part1 = ''
            try:
                venue = row[5].strip()
                if not venue:
                    raise ValueError
                timing = row[6].strip()
                if not timing:
                    raise ValueError
                part2 = '\n\nSee you at the <b>{}</b> at <b>{}</b>!'.format(
                    venue, timing)
            except:
                part2 = ''
            try:
                chairperson = row[7].strip()
                if not chairperson:
                    raise ValueError
                chairperson = '<b>Chairperson:</b> {}\n'.format(chairperson)
            except:
                chairperson = ''
            try:
                musician = row[8].strip()
                if not musician:
                    raise ValueError
                musician = '<b>Musician:</b> {}\n'.format(musician)
            except:
                musician = ''
            try:
                refreshments = row[9].strip()
                if not refreshments:
                    raise ValueError
                refreshments = '<b>Refreshments:</b> {}\n'.format(refreshments)
            except:
                refreshments = ''
            part3 = '\n\n' + chairperson + musician + refreshments
            part3 = part3.rstrip()
            try:
                ending = '\n\n' + row[10].strip()
            except:
                ending = ''
            message = intro + part1 + part2 + part3 + ending
            message = message.strip()
            return(message)
    return None


def getsheet():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])
    return values


def main():

    checker()
    schedule.every().day.at("00:00").do(checker)
    schedule.every().day.at("17:00").do(sender)

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == '__main__':
    main()

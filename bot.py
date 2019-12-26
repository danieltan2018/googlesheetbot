# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

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
from params import bottoken, channel, SPREADSHEET_ID, RANGE_NAME

bot = telegram.Bot(token=bottoken)

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.DEBUG)
logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']


def sender(message):
    bot.send_message(chat_id=channel, text=message,
                     parse_mode=telegram.ParseMode.HTML)


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
                day = row[2].strip()
                title = row[3].strip()
            except:
                continue
            try:
                speaker = row[4].strip()
                speaker = ' by <b>{}</b>'.format(speaker)
            except:
                speaker = ''
            part1 = '{} <b>{}</b>, we will be having <b>{}</b>{}.'.format(
                intro, day, title, speaker)
            try:
                venue = row[5].strip()
                timing = row[6].strip()
                part2 = '\n\nSee you at the <b>{}</b> at <b>{}</b>!'.format(
                    venue, timing)
            except:
                part2 = ''
            try:
                ending = '\n\n' + row[7].strip()
            except:
                ending = ''
            message = part1 + part2 + ending
            sender(message)


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
    print(getsheet())

    schedule.every().day.at("17:00").do(getsheet)

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == '__main__':
    main()

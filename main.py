#import statements
import urllib.request
import pandas as pd
from pushbullet import PushBullet
from zipfile import ZipFile
import os
import re
from dotenv import load_dotenv
import win32com.client as win32

from UDFs import path_check, extractZip, clean_data, organise_data

#Loading Environment Variables
load_dotenv()

#Authentication and Initialisation
pb = PushBullet(os.getenv('PUSH_BULLET_ACCESS_TOKEN'))
CHAT_DIR = os.getenv('CHAT_DIRECTORY')
EXCEL_DIR = os.getenv('SPREADSHEET_DIRECTORY')
SAVE_DIR = os.getenv("SAVES_DIRECTORY")
LAST_LINES = os.getenv('CHECK_LAST_LINES')
CHATS_TO_EXTRAPOLATE = os.getenv('CHATS_TO_EXTRAPOLATE')

#Directory existence check
path_check([CHAT_DIR, EXCEL_DIR, SAVE_DIR])

#Converting Chats_To_Extrapolate to a readable format
extrapolate_count = []
CHATS_TO_EXTRAPOLATE = CHATS_TO_EXTRAPOLATE.replace(" ", "").split("-")
if len(CHATS_TO_EXTRAPOLATE) == 1:
  extrapolate_count =  range(int(CHATS_TO_EXTRAPOLATE[0]))
else:
  extrapolate_count = range(int(CHATS_TO_EXTRAPOLATE[0])-1, int(CHATS_TO_EXTRAPOLATE[1]))

#Fetching PushBullet pushes
all_pushes = pb.get_pushes()

for current_extrapolation in extrapolate_count:
  latest_push = all_pushes[current_extrapolation]

  #Accessing required data
  file_name = latest_push['file_name']
  file_type = latest_push['file_type']
  file_url = latest_push['file_url']

  #Checking for file type and extraction
  if file_type == 'application/zip':
    urllib.request.urlretrieve(file_url, file_name)
    extractZip(f'./{file_name}', delete_zip=True)
    
  #Initialising chat data
  chat_name = file_name.split('.')[0] + '.txt'

  #Iterating through the chat
  with open(f'{CHAT_DIR}/{chat_name}', mode='r', encoding='UTF-8') as f:
    data = f.readlines()

  final_data = clean_data(data[1:])
  chat_list = organise_data(final_data)

  #Creating a dataframe
  df = pd.DataFrame(chat_list, columns=["Date", "Time", "Name", "Message"])

  #Exporting to excel
  spreadsheet_name = file_name.split('.')[0] + '.xlsx'
  df.to_excel(f"{EXCEL_DIR}/{spreadsheet_name}", index=False)

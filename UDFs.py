#Checking for file path existence
def path_check(paths):
  import os
  for path in paths:
    if not os.path.exists(f'./{path}'):
      os.makedirs(f'./{path}')

#Extracting zipfile and deleting the zip
def extractZip(file, delete_zip=False):
  from zipfile import ZipFile
  import os
  with ZipFile(file, 'r') as zObject:
    zObject.extractall(path='./All_Chats')

  if delete_zip:
    os.remove(file)

#Data Cleaner
def clean_data(uncleaned_data):
  import re
  final_data = []
  search_pattern = r'\A[0-9]+/[0-9]+/[0-9]+'
  for line_num in range(len(uncleaned_data)):
    line = uncleaned_data[line_num]
    if re.search(search_pattern, line):
      if re.search(r'\n\Z', line):
        final_data.append(line.split('\n')[0])
      else:
        final_data.append(line)
    elif re.search(r'\A\n', line):
      pass
    else:
      final_data[-1] = final_data[-1] + " " + line.split('\n')[0]
  
  return final_data

#Data Organiser
def organise_data(data):
  chat_list = []
  for line in data:
    date = line.split(',', 1)[0]
    time = line.split('-', 1)[0].split(',')[1]
    name = line.split(':', 2)[1].split('-')[1]
    message = line.split(':', 2)[2]

    chat_list.append([date, time, name, message])
  return chat_list

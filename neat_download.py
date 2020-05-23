import requests
from datetime import datetime
import time
import os
import json
import tqdm
import pywintypes, win32file, win32con

class DownloadComplete(Exception):
  pass

def log(entry):
    log_file = os.path.join(os.path.dirname(__file__), 'log', datetime.now().strftime('%m-%d-%Y_%H-%M') + '.txt')
    with open(log_file, 'a+') as f:
        f.write(entry)

def clean(text):
  unallowed = ['(', '/',  '\\', '<', '>', '"', '|', ':', '*', '?', ')']
  for ch in unallowed:
    text = text.replace(ch, '_')

  return text

def update_times(filepath, created, updated):
  ctime = datetime.strptime(created, '%Y-%m-%dT%H:%M:%SZ')
  utime = datetime.strptime(updated, '%Y-%m-%dT%H:%M:%SZ')
  modTime = time.mktime(utime.timetuple())
  os.utime(filepath, (modTime, modTime))

  wintime = pywintypes.Time(ctime)
  winfile = win32file.CreateFile(
    filepath, win32con.GENERIC_WRITE,
    win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
    None, win32con.OPEN_EXISTING,
    win32con.FILE_ATTRIBUTE_NORMAL, None)

  win32file.SetFileTime(winfile, wintime, None, None)

  winfile.close()

class Neat:
  def __init__(self, usr, pswd, dest, prev_items):
    self.usr = usr
    self.pswd = pswd
    self.dest = dest
    self.session = requests.Session()
    self.login_info = {}
    self.account_info = {}
    self.root = {}
    self.retry = False
    self.downloaded_files = prev_items

  def api_request(self, url, type, headers = None, data = None):
    self.session.headers.update(headers)
    try:
      if type == 'get':
        res = self.session.get(url)
        res = json.loads(res.text)
      else:
        res = self.session.post(url, json = data)
        res = json.loads(res.text)
      return res
    except Exception as err:
      self.retry = True
      entry = datetime.now().strftime('%m/%d/%Y %H:%M:%S') + ' - Failed complete API call: ' + url + '\n\t' + str(err) + '\n' + '*' * 80 + '\n'
      log(entry)
  
  def login(self):
    token_url = 'https://duge.neat.com/cloud/token'
    data = {
      'username': self.usr,
      'password': self.pswd
      }
    headers = {
      'User-Agent':  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36',
      'Origin': 'https://app.neat.com',
      'Referer': 'https://app.neat.com/login/'
      }
    login_res = self.api_request(token_url, 'post', headers, data)

    self.login_info = login_res

  def get_account(self):
    request_url = 'https://duge.neat.com/cloud/account'
    headers = {
      'Authorization': 'OAuth ' + self.login_info['token'],
      'Referer': 'https://app.neat.com/accounts/'
      }
    account_res = self.api_request(request_url, 'get', headers, None)

    self.account_info = account_res

  def get_root(self):
    root_url = 'https://duge.neat.com/cloud/folders/root'
    headers = {
      'x-neat-account-id': self.account_info['id'],
      'Referer': 'https://app.neat.com/dashboard/?account=' + self.account_info['id']
      }
    root_res = self.api_request(root_url, 'get', headers, None)

    self.root = root_res

  def get_folders(self, id, dest):
    subfolder_url = 'https://duge.neat.com/cloud/folders/' + id + '/subfolders'
    page = 1
    page_size = 100
    data = {
      'page': page,
      'page_size': page_size
      }
    folders_res = self.api_request(subfolder_url, 'post', {}, data)
    folder_list = folders_res['entities']
    total_folders = folders_res['pagination']['total_records']
    
    if total_folders > page_size:
      for page in range(2, total_folders + 1):
        data['page'] = page
        folders_res = self.api_request(subfolder_url, 'post', {}, data)
        folder_list.extend(folders_res['entities'])
    
    if not os.path.isdir(dest):
      os.mkdir(dest)
    
    self.get_items(id, dest)
    
    for folder in folder_list:
      if not os.path.exists(os.path.join(dest, clean(folder['name']))):
        os.mkdir(os.path.join(dest, clean(folder['name'])))
      sub_id = folder['webid']
      self.get_folders(sub_id, os.path.join(dest, clean(folder['name'])))
  
  def get_items(self, id, dest):
    items_url = 'https://duge.neat.com/cloud/items'
    page = 1
    page_size = 25
    data = {
      'filters': [{"parent_id":id},{"type":"$all_item_types"}],
      'page': page,
      'page_size': page_size,
      'utc_offset':-4
      }
    headers = {
      'Referer': 'https://app.neat.com/folders/' + self.account_info['id'] + '/?account=' + self.account_info['id']
      }
    item_res = self.api_request(items_url, 'post', headers, data)

    item_list = item_res['entities']
    total_items = item_res['pagination']['total_records']
    
    if total_items > page_size:
      for page in range(2, total_items + 1):
        data['page'] = page
        item_res = self.api_request(items_url, 'post', headers, data)
        item_list.extend(item_res['entities'])

    print('Processing folder: ' + dest)
    for item in tqdm.tqdm(item_list, desc = os.path.basename(dest)):
      if item['name'] == '':
        item_name = clean(item['created_at'][:10])
      else:
        item_name = clean(item['name'] + ' - ' + item['description'])
      item_id = item['webid']
      self.download(item_name, item_id, item['download_url'], dest, item)

  def download(self, name, id, url, dest, item):
    entry = ''
    error = ''
    filepath = os.path.join(dest, name + '.pdf')
    if id not in self.downloaded_files:
      try:
        if os.path.exists(filepath):
          ver = 1
          while os.path.exists(os.path.join(dest,name + ' (' + str(ver) + ')' + '.pdf')):
            ver += 1
          filepath = os.path.join(dest,name + ' (' + str(ver) + ')' + '.pdf')
        data = requests.get(url, timeout = 1)
        with open(filepath, 'wb') as f:
          f.write(data.content)
        update_times(filepath, item['created_at'], item['updated_at'])

        with open(os.path.join(os.path.dirname(__file__), 'prev_items.txt'), 'a+') as f:
          f.write(id + '\n')
        self.downloaded_files.append(id)

      except requests.exceptions.ConnectionError as cerr:
        error = cerr
        entry = 'Connection Error '
        self.retry = True
      except requests.exceptions.Timeout as terr:
        error = terr
        entry = 'Timeout Error '
        self.retry = True
      except requests.exceptions.HTTPError as herr:
        error = herr
        entry = 'HTTP Error '
        self.retry = True
      except requests.exceptions.RequestException as err:
        error = err
        entry = 'Unknown Error '
      
      if entry:
        entry += datetime.now().strftime('%m/%d/%Y %H:%M:%S') + ' - Failed to download file: ' + name + '\n\tSaved to: ' + dest + '\n\t' + str(error) + '\n' + '*' * 80 + '\n'
        log(entry)

def main(usr, pswd, dest):
  prev_items = []

  try:
    with open('prev_items.txt', 'r') as f:
      prev_items = f.read().split('\n')
  except FileNotFoundError as error:
    entry = datetime.now().strftime('%m/%d/%Y %H:%M:%S') + ' - prev_items.txt does not exist: ' + str(error) + '\n' + '*' * 80 + '\n'
    log(entry)

  neat_session = Neat(usr, pswd, dest, prev_items)
  neat_session.login()
  neat_session.get_account()
  neat_session.get_root()

  for folder in neat_session.root['rootFolder']['folders']:
    neat_session.get_folders(folder['webid'], os.path.join(dest, clean(folder['name'])))

  if not neat_session.retry:
    raise DownloadComplete

if __name__ == '__main__':
  if not os.path.isdir('log'):
      os.mkdir('log')
  usr = input('Enter username: ')
  pswd = input('Enter password: ')
  dest = input('Enter download path: ')
  attempt = 1
  
  while attempt < 6:
    print('Attempt', attempt)
    try:
      entry = datetime.now().strftime('%m/%d/%Y %H:%M:%S') + ' - Attempt ' + str(attempt) + '\n' + '*' * 80 + '\n'
      log(entry)
      main(usr, pswd, dest)
    except DownloadComplete:
      print('Neat file download Complete')
      # Remove or comment out line below if you would like to rerun the script without creating duplicates(assuming this is pointed to the same output folder)
      os.remove('prev_items.txt')
      break
    except Exception as error:
      entry = datetime.now().strftime('%m/%d/%Y %H:%M:%S') + ' - Unkown error: ' + str(error) + '\n' + '*' * 80 + '\n'
      log(entry)

    print('Attempt', attempt, 'did not complete successfully. Reattempting...')
    attempt += 1
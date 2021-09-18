import json
from datetime import datetime
import os


class Settings():
    def __init__(self):
        self.f = open('settings.json')
        data = json.load(self.f)
        self.oracle = data['oracle'][0]
        self.file = data['file'][0]
        self.mail = data['mail'][0]

    def get_oracle(self):
        return self.oracle

    def get_file(self):
        return self.file

    def write_debug(self, word):
        log_path = self.file['log']
        now = datetime.now().strftime('%Y%m%d')

        # Create Folder When don't have Log Fodler
        if os.path.isdir(log_path) == False:
            os.makedirs(log_path)

        log_txt = f'{now}.txt'
        log_file_path = os.path.join(log_path, log_txt)
        log = open(log_file_path, mode='a')
        string_now = datetime.now().strftime('%d/%m/%Y %X')
        log.writelines('{}--> {}\n'.format(string_now, word))
        log.close()
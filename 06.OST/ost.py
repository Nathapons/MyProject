from os import listdir, path, makedirs
from shutil import copy2
from json import load
from datetime import datetime


class Ost:
    def get_json(self):
        file = open('settings.json', )
        data = load(file)

        # Get server
        server = data['server']
        source = server[0]['source']
        destination = server[0]['destination']
        # Get File type
        about_file = data['about_file']
        file_type = about_file[0]['type']
        # Get Log
        log = data['log']
        path = log[0]['log']

        file.close()
        return source, destination, file_type, path

    def write_debug(self, log_path, word):
        now = datetime.now().strftime('%Y%m%d')

        # Create Folder When don't have Log Fodler
        if path.isdir(log_path) == False:
            makedirs(log_path)

        log_txt = f'{now}.txt'
        log_file_path = path.join(log_path, log_txt)
        log = open(log_file_path, mode='a')
        string_now = datetime.now().strftime('%d/%m/%Y %X')
        log.writelines('{}--> {}\n'.format(string_now, word))
        log.close()

    def copy_file(self, **params):
        self.src = params['source']
        self.des = params['destination']
        self.file_type = params['file_type']
        file_list = self.file_type.split(',')
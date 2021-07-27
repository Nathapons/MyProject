from xml.dom import minidom
from datetime import datetime
import os


class Settings:
    def __init__(self):
        doc = minidom.parse('./AppConfig.xml')
        add = doc.getElementsByTagName('add')

        self.log_path = add[0].getAttribute('value')
        self.yamaha = add[1].getAttribute('value')
        self.nidech = add[2].getAttribute('value')
        self.destination = add[3].getAttribute('value')
        self.ort_path = add[4].getAttribute('value')
        not_read = add[5].getAttribute('value')
        self.not_read_list = not_read.split(', ')

    def write_debug(self, word):
        now = datetime.now().strftime('%Y%m%d')

        if os.path.isdir(self.log_path) == False:
            os.makedirs(self.log_path)

        log_txt = f'{now}.txt'
        log_file_path = os.path.join(self.log_path, log_txt)
        log = open(log_file_path, mode='a')
        str_now = datetime.now().strftime('%d/%m/%Y %X')
        log.writelines(f'{str_now}--> {word}\n')
        log.close()


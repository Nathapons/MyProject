from datetime import datetime
from xml.dom import minidom
import cx_Oracle
import os
import win32com.client


class Settings:
    def connect_oracle(self):
        oracle, config = self.read_config()
        dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
        conn = cx_Oracle.connect(user=oracle['user'], password=oracle['password'], dsn=dsn_tns)
        return conn

    def disconnect_oracle(self, conn):
        conn.close()
    
    def write_debug(self, word, log_path):
        now = datetime.now().strftime('%Y%m%d')

        # Create Folder When don't have Log Fodler
        if os.path.isdir(log_path) == False:
            os.makedirs(log_path)

        now_file = os.path.join(log_path, now)
        log_txt = f'{now}.txt'
        log_file_path = os.path.join(log_path, log_txt)
        log = open(log_file_path, mode='a')
        string_now = datetime.now().strftime('%d/%m/%Y %X')
        log.writelines('{}--> {}\n'.format(string_now, word))
        log.close()

    def read_config(self):
        path = os.path.dirname(os.path.abspath(__file__))
        xml_file = "AppConfig.xml"
        doc = minidom.parse(xml_file)

        # Start read XML file
        sql = doc.getElementsByTagName('sql')
        # Get Database Information
        client = sql[0].getAttribute('value')
        user = sql[1].getAttribute('value')
        password = sql[2].getAttribute('value')

        oracle = {
            'client': client,
            'user': user,
            'password': password
        }
        # Get App Settings
        add = doc.getElementsByTagName('add')
        alert_to = add[0].getAttribute('value')
        alert_cc = add[1].getAttribute('value')
        log = add[2].getAttribute('value')
        destination = add[3].getAttribute('value')
        source = add[4].getAttribute('value')
        file_type = add[5].getAttribute('value')
        config = {
            'alert_to': alert_to,
            'alert_cc': alert_cc,
            'log': log,
            'destination': destination,
            'source': source,
            'file_type': file_type
        }
        
        return oracle, config

    def send_mail(self, attach_file=None, html_body=None):
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        # Mail Detail
        mail.To = self.config['TO']
        mail.CC = self.config['CC']
        mail.Subject = 'Python Test mail'
        mail.HTMLBody = html_body

        # Attach file to user
        if attach_file is not None:
            mail.Attachments.Add(attach_file)

        # Send mail to User
        mail.Send()


# if __name__ == "__main__":
#     app = Settings()
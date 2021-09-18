from datetime import datetime
import shutil
from xml.dom import minidom
import cx_Oracle
import pandas as pd
from os import mkdir, path, environ, remove
from shutil import move
import win32com.client
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from json import load


class Settings:
    def __init__(self):
        file = open(file='config.json')
        data = load(file)
        self.oracle = data['oracle'][0]
        self.file = data['file'][0]
        self.mail = data['mail'][0]
        file.close()

    def get_oracle(self):
        return self.oracle

    def get_file(self):
        return self.file

    def get_mail(self):
        return self.mail

    def write_debug(self, word):
        log_path = self.file['log']
        now = datetime.now().strftime('%Y%m%d')

        # Create Folder When don't have Log Fodler
        if path.isdir(log_path) == False:
            mkdir(log_path)

        log_txt = f'{now}.txt'
        log_file_path = path.join(log_path, log_txt)
        log = open(log_file_path, mode='a')
        string_now = datetime.now().strftime('%d/%m/%Y %X')
        log.writelines('{}--> {}\n'.format(string_now, word))
        log.close()

    def send_mail_win32(self, subject, body, attach_file=None):
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = self.mail_to
        mail.CC = self.mail_cc
        mail.Subject = subject
        mail.HTMLBody = body

        if attach_file is not None:
            mail.Attachments.Add(attach_file)

        mail.Send()

    def send_mail(self, body):
        try:
            message = MIMEMultipart("alternative")
            message["Subject"] = f"Alert Test email: {datetime.now().strftime('%d/%m/%Y %X')}"
            message["From"] = "iot_mail@th.fujikura.com"
            message["To"] = self.mail_to
            message["CC"] = self.mail_cc
            part = MIMEText(body, "html")

            message.attach(part)
            smtObj = smtplib.SMTP("116.68.146.43")
            smtObj.sendmail(from_addr="mail@th.fujikura.com", to_addrs=self.mail_cc, msg=message.as_string())
        except Exception as error:
            self.write_debug(word="Unable send mail")
        finally:
            self.write_debug(word="Send mail complete")

    def send_mail_html(self, datatable):
        # Datatable => dict in list => [{}]
        html_body = ""

        if len(datatable) > 0:
            # Create Table Title
            html_body += "<p style='font-size: 20px; font-weight: normal'>Defect Report from IOT<p>"

            # Create Table Mail
            html_body += "	    <table cellpadding='0' cellspacing='0' width='700px'> \n"
            html_body += "	        <tr style='height: 25px'>\n"
            html_body += "	            <td style='text-align: center; width: 80px; background-color: Red; color: White; border:1px solid gray'>No.</td>\n"
            html_body += "	            <td style='text-align: center; width: 80px; background-color: Red; color: White; border:1px solid gray'>A</td>\n"
            html_body += "	            <td style='text-align: center; width: 80px; background-color: Red; color: White; border:1px solid gray'>B</td>\n"
            html_body += "	            <td style='text-align: center; width: 80px; background-color: Red; color: White; border:1px solid gray'>C</td>\n"
            html_body += "	        </tr>"

            row_no = 1
            for row in datatable:
                html_body += "	        <tr style='height: 25px'>\n"
                html_body += f"	            <td style='text-align: center; border:1px solid gray'>{row_no}</td>\n"
                html_body += f"	            <td style='text-align: center; border:1px solid gray'>{row['a']}</td>\n"
                html_body += f"	            <td style='text-align: center; border:1px solid gray'>{row['b']}</td>\n"
                html_body += f"	            <td style='text-align: center; border:1px solid gray'>{row['c']}</td>\n"
                html_body += "	        </tr>"
                row_no += 1
            html_body += "	    </table>\n"

            html_body += "<p style='font-size: 15px; font-weight: normal'>Remark: Can't reply this email</p>\n"
            html_body += """
                <div>
                    <p style='font-size: 15px; font-weight: normal'>----------------------------------------</p>
                    <p style='font-size: 15px; font-weight: normal'>Fujikura Electronics (Thailand) Ltd.</p>
                    <p style='font-size: 15px; font-weight: normal'>Ayutthaya Factory(A1)</p>
                    <p style='font-size: 15px; font-weight: normal'>Internet of things division</p>
                    <p style='font-size: 15px; font-weight: normal'>Tel. 4308</p>
                </div>
            """

            # Run function send mail to user
            self.send_mail(body=html_body)

    def connect_fpc(self):
        environ['PATH'] = self.client
        dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
        username = self.oracle.get('fpc_username')
        password = self.oracle.get('fpc_password')
        conn = cx_Oracle.connect(user=username, password=password, dsn=dsn_tns)
        return conn

    def select_fpc(self, sql_command):
        conn = self.connect_fpc()
        datatable = None

        try:
            datatable = pd.read_sql(sql_command, con=conn)
        except Exception as error:
            self.write_debug(word=f"  Error : {str(error)}")
        finally:
            conn.close()

        return datatable

    def save_fpc(self, sql_command):
        conn = self.connect_fpc()
        str_error = ""

        try:
            cur = conn.cursor()
            cur.execute(sql_command)
            conn.commit()
        except Exception as error:
            str_error = str(error)
            self.write_debug(word=f"  Error : {str(error)}")
        finally:
            conn.close()

        return str_error

    def connect_irpt(self):
        client = self.oracle.get('client')
        environ['PATH'] = client
        username = self.oracle.get('irpt_username')
        password = self.oracle.get('irpt_password')
        dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
        conn = cx_Oracle.connect(user=username, password=password, dsn=dsn_tns)
        return conn

    def select_irpt(self, sql_command):
        conn = self.connect_irpt()
        datatable = None

        try:
            datatable = pd.read_sql(sql_command, con=conn)
        except Exception as error:
            self.write_debug(word=f"  Error : {str(error)}")
        finally:
            conn.close()

        return datatable

    def save_irpt(self, sql_command):
        conn = self.connect_irpt()
        str_error = ""

        try:
            cur = conn.cursor()    
            cur.execute(sql_command)
            conn.commit()
        except Exception as error:
            str_error = str(error)
        finally:
            conn.close()

        return str_error

    def move_file(self, src, dst):
        if path.isfile(dst):
            remove(dst)
        shutil.move(src=src, dst=dst)
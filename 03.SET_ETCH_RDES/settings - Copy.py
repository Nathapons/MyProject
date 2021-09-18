from datetime import datetime
from xml.dom import minidom
import cx_Oracle
import pandas as pd
import os 
import win32com.client
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class Settings:
    def __init__(self):
        doc = minidom.parse("D:\\Nathapon\\0.My work\\01.IoT\\00.AX\\PYTHON\\03.SET_ETCH_RDES\\AppConfig.xml")

        # Start read XML file
        add = doc.getElementsByTagName('add')
        # Get Database Information
        self.dsn = add[0].getAttribute('value')
        self.client = add[1].getAttribute('value')
        self.report_user = add[2].getAttribute('value')
        self.report_password = add[3].getAttribute('value')
        self.fpc_user = add[4].getAttribute('value')
        self.fpc_password = add[5].getAttribute('value')
        self.mail_to = add[6].getAttribute('value')
        self.mail_cc = add[7].getAttribute('value')
        self.source_path = add[8].getAttribute('value')
        self.out_path = add[9].getAttribute('value')
        self.error_path = add[10].getAttribute('value')
        self.log_path = add[11].getAttribute('value')
        self.file_type = add[12].getAttribute('value')
    
    def write_debug(self, word):
        now = datetime.now().strftime('%Y%m%d')

        # Create Folder When don't have Log Fodler
        if os.path.isdir(self.log_path) == False:
            os.makedirs(self.log_path)

        now_file = os.path.join(self.log_path, now)
        log_txt = f'{now}.txt'
        log_file_path = os.path.join(self.log_path, log_txt)
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
        os.environ['PATH'] = self.client
        dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
        conn = cx_Oracle.connect(user=self.fpc_user, password=self.fpc_password, dsn=dsn_tns)
        return conn

    def select_fpc(self, sql_command):
        conn = self.connect_fpc()
        datatable = None

        try:
            datatable = pd.read_sql(sql_command, con=conn)
            # cur = conn.cursor()    
            # cur.execute(sql_command)
            # datatable = cur.fetchall()
        except Exception as error:
            self.write_debug(word=f"  Error : {str(error)}")
        finally:
            conn.close()

        return datatable

    def save_fpc(self, sql_command):
        conn = self.connect_fpc()
        str_error = ""

        try:
            # conn.autocommit = True
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
        os.environ['PATH'] = self.client
        dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
        conn = cx_Oracle.connect(user=self.report_user, password=self.report_password, dsn=dsn_tns)
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

    def get_max_header(self, header_code, factory, str_date, str_time, str_mc):
        sql_command = f"""
            SELECT MAX(T.SRH_CODE) AS HC
            FROM SMF_RECORD_HEADER T
            WHERE T.SRH_CODE LIKE '{header_code}%'
                  AND T.SRH_GROUP = '0001'
                  AND T.SRH_UNIT = 'CFM'
                  AND T.SRH_MC = '{str_mc}'
                  AND T.SRH_KEY_3 = '{factory}'
                  AND T.SRH_KEY_4 = '{str_date}'
                  AND T.SRH_KEY_5 = '{str_time}'
        """
        hc_df = self.select_irpt(sql_command=sql_command)
        max_hc = hc_df.loc[0, 'HC']
        if max_hc is not None:
            return max_hc

        sql_command = f"""
            SELECT MAX(T.SRH_CODE) AS HC
            FROM SMF_RECORD_HEADER T
            WHERE T.SRH_CODE LIKE '{header_code}%'
                  AND T.SRH_UNIT = 'CFM'
                  AND T.SRH_GROUP = '0001'
        """
        hc_df = self.select_irpt(sql_command=sql_command)
        max_hc = hc_df.loc[0, 'HC']
        if max_hc is not None:
            return int(max_hc) + 1

        max_hc = header_code + "0001"
        return max_hc   


if __name__ == "__main__":
    app = Settings()
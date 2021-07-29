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
        doc = minidom.parse("./AppConfig.xml")

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
        dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
        conn = cx_Oracle.connect(user=self.report_user, password=self.report_password, dsn=dsn_tns)
        return conn

    def select_irpt(self, sql_command):
        conn = self.connect_irpt()
        datatable = None

        try:
            cur = conn.cursor()    
            cur.execute(sql_command)
            datatable = cur.fetchall()
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
            self.write_debug(word=f"  Error : {str(error)}")
        finally:
            conn.close()

        return str_error

    def insert_data(self, str_filename, df, file_code, df_code):
        str_error = ''
        sql_commands = []
        total_row = df.shape[0]

        if total_row > 0:
            for index, df_row in df.iterrows():
                date = df_row['23']
                time = df_row['2']
                machine = df_row['24']
                lot_no = df_row['25']

                sql_header =  "   MERGE INTO FPCC_WORKING_RECORD_HEADER D\n"
                sql_header += f"   USING (SELECT TO_DATE('{date} {time}','DD/MM/YYYY HH24:MI') AS WRH_DATE\n"
                sql_header += f"               , '{machine}' AS WRH_MACHINE\n"
                sql_header += f"               , '{lot_no}' AS WRH_LOT_NO\n"
                sql_header += "               , '-' AS WRH_SHIFT\n"
                sql_header += "               , '-' AS WRH_CHECK_BY\n"
                sql_header += "               , '-' AS WRH_PRD_TYPE\n"
                sql_header += "               , 'CFM' AS WRH_PROCESS_ID\n"
                sql_header += "               , '-' AS WRH_PATH\n"
                sql_header += f"               , TO_DATE('{date} {time}','DD/MM/YYYY HH24:MI') AS WRH_CHECK_DATE\n"
                sql_header += f"               , '{file_code}' AS WRH_TYPE\n"
                sql_header += "               , 'Y' AS WRH_STATUS\n"
                sql_header += f"               , '{str_filename}' AS WRH_FILE_NAME\n"
                sql_header += "               , (SELECT (NVL(MAX(WRT.WRT_REV),1)) AS MAX_REV FROM FPCC_WORKING_RECORD_TYPE WRT WHERE WRT.WRT_TYPE='00001') AS WRH_REV\n"
                sql_header += "                            FROM DUAL\n"
                sql_header += "                            ) R\n"
                sql_header += "                   ON (\n"
                sql_header += "                          D.WRH_DATE = R.WRH_DATE\n"
                sql_header += "                      AND D.WRH_TYPE = R.WRH_TYPE\n"
                sql_header += "                      AND D.WRH_MACHINE = R.WRH_MACHINE\n"
                sql_header += "                      AND D.WRH_LOT_NO = R.WRH_LOT_NO\n"
                sql_header += "                      AND D.WRH_REV = R.WRH_REV\n"
                sql_header += "                   )\n"
                sql_header += "                   WHEN MATCHED THEN\n"
                sql_header += "                     UPDATE\n"
                sql_header += "                        SET D.WRH_STATUS = R.WRH_STATUS ,\n"
                sql_header += "                            D.WRH_SHIFT = R.WRH_SHIFT ,\n"
                sql_header += "                            D.WRH_CHECK_BY = D.WRH_CHECK_BY ,\n"
                sql_header += "                            D.WRH_CHECK_DATE = D.WRH_CHECK_DATE,\n"
                sql_header += "                            D.WRH_FILE_NAME = R.WRH_FILE_NAME,\n"
                sql_header += "                            D.WRH_PROCESS_ID = R.WRH_PROCESS_ID,\n"
                sql_header += "                            D.WRH_PRD_TYPE = R.WRH_PRD_TYPE,\n"
                sql_header += "                            D.WRH_PATH = R.WRH_PATH\n"
                sql_header += "                   WHEN NOT MATCHED THEN\n"
                sql_header += "                     INSERT (D.WRH_DATE\n"
                sql_header += "                             , D.WRH_TYPE\n"
                sql_header += "                             , D.WRH_MACHINE\n"
                sql_header += "                             , D.WRH_LOT_NO\n"
                sql_header += "                             , D.WRH_STATUS\n"
                sql_header += "                             , D.WRH_SHIFT\n"
                sql_header += "                             , D.WRH_CHECK_BY\n"
                sql_header += "                             , D.WRH_CHECK_DATE\n"
                sql_header += "                             , D.WRH_FILE_NAME\n"
                sql_header += "                             , D.WRH_PRD_TYPE\n"
                sql_header += "                             , D.WRH_PROCESS_ID\n"
                sql_header += "                             , D.WRH_PATH\n"
                sql_header += "                             , D.WRH_REV\n"
                sql_header += "                             )\n"
                sql_header += "                    VALUES(R.WRH_DATE\n"
                sql_header += "                             , R.WRH_TYPE\n"
                sql_header += "                             , R.WRH_MACHINE\n"
                sql_header += "                             , R.WRH_LOT_NO\n"
                sql_header += "                             , R.WRH_STATUS\n"
                sql_header += "                             , R.WRH_SHIFT\n"
                sql_header += "                             , R.WRH_CHECK_BY\n"
                sql_header += "                             , R.WRH_CHECK_DATE\n"
                sql_header += "                             , R.WRH_FILE_NAME\n"
                sql_header += "                             , R.WRH_PRD_TYPE\n"
                sql_header += "                             , R.WRH_PROCESS_ID\n"
                sql_header += "                             , R.WRH_PATH\n"
                sql_header += "                             , R.WRH_REV\n"
                sql_header += "                             )"
                sql_commands.append(sql_header)

                i = 1
                for index, row in df_code.iterrows():
                    wrm_code = row['CODE']
                    wrd_value = df_row[str(i)]

                    sql_detail = " MERGE INTO FPCC_WORKING_RECORD_DETAIL D\n"
                    sql_detail += f" USING (SELECT TO_DATE('{date} {time}','DD/MM/YYYY HH24:MI') AS WRD_DATE\n"
                    sql_detail += f"             , '{machine}' AS WRD_MACHINE\n"
                    sql_detail += f"             , '{lot_no}' AS WRD_LOT_NO\n"
                    sql_detail += f"             , '{file_code}' AS WRD_TYPE\n"
                    sql_detail += f"             , '{wrm_code}' AS WRD_CODE\n"
                    sql_detail += f"             , '{wrd_value}' AS WRD_VALUE\n"
                    sql_detail += "             , 'A' AS WRD_STATUS\n"
                    sql_detail += "             , '1' AS WRD_SEQ\n"
                    sql_detail += "             , (SELECT (NVL(MAX(WRT.WRT_REV),1)) AS MAX_REV FROM FPCC_WORKING_RECORD_TYPE WRT WHERE WRT.WRT_TYPE='00001') AS WRD_REV\n"
                    sql_detail += "                          FROM DUAL\n"
                    sql_detail += "                          ) R\n"
                    sql_detail += "                 ON (\n"
                    sql_detail += "                    D.WRD_DATE = R.WRD_DATE and\n"
                    sql_detail += "                    D.WRD_TYPE = R.WRD_TYPE and\n"
                    sql_detail += "                    D.WRD_MACHINE = R.WRD_MACHINE and\n"
                    sql_detail += "                    D.WRD_LOT_NO = R.WRD_LOT_NO and\n"
                    sql_detail += "                    D.WRD_LOT_NO = R.WRD_LOT_NO and\n"
                    sql_detail += "                    D.WRD_CODE = R.WRD_CODE and\n"
                    sql_detail += "                    D.WRD_REV = R.WRD_REV and\n"
                    sql_detail += "                    D.WRD_SEQ = R.WRD_SEQ\n"
                    sql_detail += "                 )\n"
                    sql_detail += "                 WHEN MATCHED THEN\n"
                    sql_detail += "                   UPDATE\n"
                    sql_detail += "                      SET D.WRD_VALUE = R.WRD_VALUE\n"
                    sql_detail += "                 WHEN NOT MATCHED THEN\n"
                    sql_detail += "                   INSERT (D.WRD_DATE\n"
                    sql_detail += "                           , D.WRD_TYPE\n"
                    sql_detail += "                           , D.WRD_MACHINE\n"
                    sql_detail += "                           , D.WRD_LOT_NO\n"
                    sql_detail += "                           , D.WRD_CODE\n"
                    sql_detail += "                           , D.WRD_VALUE\n"
                    sql_detail += "                           , D.WRD_STATUS\n"
                    sql_detail += "                           , D.WRD_REV\n"
                    sql_detail += "                           , D.WRD_SEQ\n"
                    sql_detail += "                           )\n"
                    sql_detail += "                   VALUES(R.WRD_DATE\n"
                    sql_detail += "                           , R.WRD_TYPE\n"
                    sql_detail += "                           , R.WRD_MACHINE\n"
                    sql_detail += "                           , R.WRD_LOT_NO\n"
                    sql_detail += "                           , R.WRD_CODE\n"
                    sql_detail += "                           , R.WRD_VALUE\n"
                    sql_detail += "                           , R.WRD_STATUS\n"
                    sql_detail += "                           , R.WRD_REV\n"
                    sql_detail += "                           , R.WRD_SEQ\n"
                    sql_detail += "                           )"

                    sql_commands.append(sql_detail)
                    i += 1

            for sql_command in sql_commands:
                self.save_fpc(sql_command=sql_command)
    

if __name__ == "__main__":
    app = Settings()
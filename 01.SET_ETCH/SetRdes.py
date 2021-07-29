from settings import Settings
import os
import shutil
import datetime
import pandas as pd


class SetRdes:
    def __init__(self):
        self.config = Settings()
        self.config.write_debug(word="Start Program")
        self.get_file_dir()
        self.config.write_debug(word="End Program")

    def get_file_dir(self):
        # Get filename in Folder IN
        file_dir = os.listdir(self.config.source_path)

        if len(file_dir) > 0:
            for file in file_dir:
                now = datetime.datetime.now()
                now_format = now.strftime('%Y%m%d_%H%m_')
                new_file = now_format + file
                file_in = os.path.join(self.config.source_path, file)
                file_out = os.path.join(self.config.out_path, new_file)
                file_error = os.path.join(self.config.error_path, new_file)

                try:
                    if str(file).endswith(self.config.file_type) and (not str(file).startswith('~$')):
                        self.config.write_debug(word=f"     Open file : {str(file)}")
                        self.open_excel_file(file_path=file_in)

                        if os.path.isfile(file_out):
                            os.remove(file_out) 
                        shutil.move(file_in, file_out)
                except Exception as e:
                    self.config.write_debug(word=f"     Error=> {e}")
                    if os.path.isfile(file_error):
                        os.remove(file_error)
                    shutil.move(file_in, file_error)
                finally:
                    self.config.write_debug(word=f"     End Read file : {str(file)}")
                    
        else:
            self.config.write_debug(word="     There is no FILE. Wait until next event")

    def open_excel_file(self, file_path):
        file = pd.ExcelFile(file_path)
        sheetnames = [sheet for sheet in file.sheet_names if 'DES' not in sheet.upper()]
        # sheetnames = file.sheet_names
        file.close()

        dt_raw = []
        dt_etch_alarm = []
        dt_chem_alarm = []
        for sheetname in sheetnames:
            self.config.write_debug(word=f"     Read Excel at sheetname: {sheetname}")
            df = pd.read_excel(file_path, sheetname, header=9)
            
            count = 0
            for i in range(len(df)):
                date = df.iloc[i, 0]
                if str(date) != 'nan':
                    count += 1

            if count == 0:
                continue

            for i in range(len(df)):
                try:
                    title = df.iloc[i, 4]
                    if 'ETCH' in title.upper():
                        new_row = {}
                        time = df.iloc[i-2,0]
                        date = df.iloc[i-3,0]
                        if date is nan and time is nan:
                            break
                        
                        str_date = ""
                        str_time = ""
                        try:
                            if '00:00:00' in str(date):
                                filename = os.path.basename(file_path)
                                date_full = str(date).split()[0]
                                day = date_full.split('-')[1]
                                month = str(filename).split('_')[1][:2]
                                year = date_full.split('-')[0]
                                str_date = day + "/" + month + "/" + year
                            else:
                                date_split = str(date).split('-')
                                str_date = date_split[0] + '/' + date_split[1] + '/' + date_split[2]

                            if '-' in str(time):
                                time = str(time).split(' ')[1]

                            time_list = str(time).split(':')
                            hour = time_list[0]
                            mins = time_list[1]
                            str_time = hour + ':' + mins
                        except Exception as error:
                            str_date = nan
                            str_time = nan
                        
                        # Upper Etching
                        wrm_code = 6
                        upper_empty = False
                        upper_list = []
                        for col in range(5, 10):
                            str_code = str(wrm_code)
                            weight_before = df.iloc[i-3, col]
                            weight_after = df.iloc[i-2, col]
                            upper_val = df.iloc[i, col]

                            try:
                                if weight_after is nan or weight_before is nan and (not isinstance(upper_val, float)):
                                    upper_empty = True
                                else:
                                    upper_val = round(upper_val, 3)
                                    new_row[str_code] = upper_val
                                    upper_list.append(upper_val)
                                    wrm_code += 1
                            except Exception as error:
                                upper_empty = True
                        
                        # Lower Etching
                        wrm_code = 12
                        lower_empty = False
                        lower_list = []
                        for col in range(13, 18):
                            str_code = str(wrm_code)
                            weight_before = df.iloc[i-3, col]
                            weight_after = df.iloc[i-2, col]
                            lower_val = df.iloc[i, col]

                            try:
                                if weight_after is nan or weight_before is nan and (not isinstance(lower_val, float)):
                                    lower_empty = True
                                else:
                                    lower_val = round(lower_val, 3)
                                    new_row[str_code] = lower_val
                                    lower_list.append(lower_val)
                                    wrm_code += 1
                            except Exception as error:
                                lower_empty = True

                        # Uniformity
                        upper_uni = round(df.iloc[i, 11], 3)
                        new_row['11'] = upper_uni
                        lower_uni = round(df.iloc[i, 19], 3)
                        new_row['17'] = lower_uni

                        # Chemical
                        wrm_code = 18
                        chem_empty = False
                        chem_list = []
                        for col in range(21, 24):
                            chemical = df.iloc[i, col]
                            try:
                                if chemical is nan and (not isinstance(chemical, float)):
                                    chem_empty = True
                                else:
                                    chemical = round(df.iloc[i, col], 3)
                                    str_code = str(wrm_code)
                                    new_row[str_code] = chemical
                                    chem_list.append(chemical)
                                    wrm_code += 1
                            except Exception as error:
                                chem_empty = True

                        if upper_empty or lower_empty or chem_empty:
                            continue

                        building = os.path.basename(file_path).split('_')[0]
                        machine = str(sheetname).strip()
                        target = df.iloc[i-3, 1]
                        line_speed = df.iloc[i-3, 2]
                        line = df.iloc[i-3, 3]
                        if str(line) == 'nan':
                            line = '-'
                        record = df.iloc[i, 24]

                        new_row['1'] = building
                        new_row['2'] = str_time
                        new_row['3'] = target
                        new_row['4'] = line_speed
                        new_row['5'] = (line if line is not nan else '-')
                        new_row['21'] = record
                        new_row['22'] = '-'
                        new_row['23'] = str_date
                        new_row['24'] = machine
                        new_row['25'] = '-'

                        if nan not in new_row.values():
                            dt_raw.append(new_row)

                except Exception as error:
                    pass

        df = pd.DataFrame(dt_raw)

        # if len(df) > 0:    
        #     # Insert Datatable
        #     file_code = '00048'
        #     df_code = self.get_master_by_wrm_code(file_code)
        #     filename = os.path.basename(file_path)
        #     self.config.write_debug(word="     Start insert to Datatabase")
        #     self.config.insert_data(str_filename=filename, df=df, file_code=file_code, df_code=df_code)
        #     self.config.write_debug(word="     End insert to Datatabase")

    def get_master_by_wrm_code(self, file_code):
        sql_command = f"""
            SELECT T.WRM_CODE AS CODE
            FROM FPCC_WORKING_RECORD_MASTER T
            WHERE T.WRM_TYPE = '{file_code}'
        """
        df_code = self.config.select_fpc(sql_command=sql_command)
        return df_code

    def get_df_spec(self, machine, measure_type):
        sql_command = f"""
            SELECT   
                      T.WPM_PRODUCT AS MACHINE
                    , T.WPM_LCL AS LCL
                    , T.WPM_UCL AS UCL
            FROM FPCC_WORKING_PRODUCT_MASTER T
            WHERE 
                      T.WPM_PRODUCT = '{machine}'
                  AND T.WPM_KEY_1 = '{measure_type}'
        """
        df_spec = self.config.select_fpc(sql_command=sql_command)
        return df_spec
           

if __name__ == '__main__':
    app = SetRdes()
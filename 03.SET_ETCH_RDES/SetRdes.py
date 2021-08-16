from settings import Settings
import os
import shutil
import datetime
import pandas as pd


class SetRdes:
    def __init__(self):
        self.total_commit = 0
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
                        # self.etch_program(file_path=file_in)
                        self.des_program(file_path=file_in)

                        # if os.path.isfile(file_out):
                        #     os.remove(file_out)
                        # shutil.move(file_in, file_out)
                except Exception as e:
                    self.config.write_debug(word=f"     Error=> {e}")
                    # if os.path.isfile(file_error):
                    #     os.remove(file_error)
                    # shutil.move(file_in, file_error)
                finally:
                    self.config.write_debug(word=f"     End Read file : {str(file)}")
                    
        else:
            self.config.write_debug(word="     There is no FILE. Wait until next event")

    def etch_program(self, file_path):
        file = pd.ExcelFile(file_path)
        sheetnames = [sheet for sheet in file.sheet_names if 'DES' not in sheet.upper()]
        file.close()

        dt_raw = []
        for sheetname in sheetnames:
            self.config.write_debug(word=f"     ETCH Program read Excel at sheetname: {sheetname}")
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
                        new_row['5'] = (line if str(line) != 'nan' else '-')
                        new_row['21'] = record
                        new_row['22'] = '-'
                        new_row['23'] = str_date
                        new_row['24'] = machine
                        new_row['25'] = '-'

                        if 'nan' not in new_row.values():
                            dt_raw.append(new_row)

                except Exception as error:
                    pass

        df = pd.DataFrame(dt_raw)

        if len(df) > 0:    
            # Insert Datatable
            file_code = '00048'
            df_code = self.get_master_by_wrm_code(file_code)
            filename = os.path.basename(file_path)
            self.config.write_debug(word="     Start insert to ETCH Datatabase")
            self.config.insert_working_record(str_filename=filename, df=df, file_code=file_code, df_code=df_code)
            self.config.write_debug(word="     End insert to ETCH Datatabase")

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

    def des_program(self, file_path):
        filename = os.path.basename(file_path)
        file = pd.ExcelFile(file_path)
        sheetnames = [sheet for sheet in file.sheet_names if 'DES' in sheet.upper()]
        file.close()

        datatable = []
        for sheetname in sheetnames:
            df = pd.read_excel(file_path, sheetname, header=8)
            if len(df) == 0:
                continue
            
            for i in range(len(df)):
                weight_title = df.loc[i, 'น้ำหนัก']

                if 'ETCH' in str(weight_title).upper():
                    datarow = {}
                    date = df.iloc[i-3, 0]
                    time = df.iloc[i-2, 0]
                    date_form = ""
                    header_date = ""
                    time_form = ""
                    header_time = ""

                    if str(date) == 'nan' or str(time) == 'nan':
                        continue

                    try:
                        if isinstance(date, datetime.datetime):
                            date_form = date.strftime('%d/%m/%Y')
                            date_sql = date.strftime('%Y%m%d')

                        if isinstance(date, str):
                            date_list = str(date).split('-')
                            date_form = date_list[0] + "/" + date_list[1] + "/" + date_list[2]
                            date_sql = date_list[2] + date_list[1] + date_list[0]

                        time_list = str(time).split(':')
                        hrs = time_list[0]
                        mins = time_list[1]
                        time_form = hrs + "." + mins
                    except Exception as error:
                        str_error = str(error)
                        # self.config.write_debug(word=f"      Convert at row {i+10} has error {str_error}")
                        continue 

                    # Etching Upper
                    has_upper = True
                    seq = 4
                    col = 5
                    while col <= 9:
                        weight_before = df.iloc[i-3, col]
                        weight_after = df.iloc[i-2, col]
                        etch_upper = ""

                        if str(weight_before) == 'nan' and str(weight_after) == 'nan':
                            has_upper = False
                        else:
                            etch_upper = df.iloc[i, col]
                            str_seq = str(seq)
                            datarow[str_seq] = round(etch_upper, 2)

                        seq += 1
                        col += 1
                    
                    # Etching Lower
                    has_lower = True
                    seq = 10
                    col = 13
                    while col <= 17:
                        weight_before = df.iloc[i-3, col]
                        weight_after = df.iloc[i-2, col]
                        etch_lower = ""

                        if str(weight_before) == 'nan' and str(weight_after) == 'nan':
                            has_lower = False
                        else:
                            etch_lower = df.iloc[i, col]
                            str_seq = str(seq)
                            datarow[str_seq] = round(etch_lower, 2)
                        
                        seq += 1
                        col += 1

                    # Chemical centration
                    has_chem = True
                    col = 21
                    seq = 16
                    while col <= 23:
                        chem = ''

                        if str(chem) == 'nan':
                            has_chem = False
                        else:
                            chem = df.iloc[i, col]
                        # Append Datarow
                        str_seq = str(seq)
                        datarow[str_seq] = round(chem, 2)

                        seq += 1
                        col += 1

                    if has_upper and has_lower and has_chem:
                        target = df.iloc[i-3, 1]
                        datarow['1'] = target
                        line_speed = df.iloc[i-3, 2]
                        datarow['2'] = line_speed
                        line = df.iloc[i-3, 3]
                        datarow['3'] = (line if str(line) != 'nan' else '-')
                        upper_uni = df.iloc[i, 11]
                        datarow['9'] = round(upper_uni, 2)
                        lower_uni = df.iloc[i, 19]
                        datarow['15'] = round(lower_uni, 2)
                        
                        datarow['date'] = date_form
                        datarow['time'] = time_form
                        datarow['date_sql'] = date_sql
                        operator = df.iloc[i, 24]
                        factory = filename.split('_')[0]
                        datarow['factory'] = factory
                        datarow['operator'] = operator
                        datarow['machine'] = sheetname
                        datatable.append(datarow)

        rdes = pd.DataFrame(datatable)
        if len(rdes) > 0:
            smf_group = '0001'
            self.config.write_debug(word="     Start insert to RDES Datatabase")
            self.insert_smf_record(str_filename=filename, df=rdes, smf_group=smf_group)   
            self.config.write_debug(word="     End insert to RDES Datatabase")

    def insert_smf_record(self, str_filename, df, smf_group):
        query_list = []
        
        for i in range(len(df)):
            date = df.loc[i, 'date']
            machine = df.loc[i, 'machine']
            date_sql = df.loc[i, 'date_sql']
            factory = df.loc[i, 'factory']
            operator = df.loc[i, 'operator']
            time = df.loc[i, 'time']
            code = self.config.get_max_header(
                header_code=date_sql,
                factory=factory,
                str_date=date,
                str_time=time,
                str_mc=machine
            )
            
            if str(code) == "" or str(time) == "" or str(date) == "":
                continue
        
            header = f"""
                MERGE INTO SMF_RECORD_HEADER T
                USING (
                    SELECT
                            'CFM' AS SRD_UNIT
                            , '{smf_group}' AS SRD_GROUP
                            , '{code}' AS SRD_CODE
                            , TO_DATE('{date}', 'DD/MM/YYYY') AS SRD_DATE
                            , '-' AS SRD_LOT_NO
                            , '-' AS SRD_PROCESS
                            , '{machine}' AS SRD_MC
                            , '-' AS SRD_KEY_1
                            , '-' AS SRD_KEY_2
                            , '{factory}' AS SRD_KEY_3
                            , '{date}' AS SRD_KEY_4
                            , '{time}' AS SRD_KEY_5
                            , '{operator}' AS SRD_OPERATOR
                            , '{time}' AS SRD_TIME
                            , '{str_filename}' AS SRD_FILENAME
                    FROM DUAL
                ) D
                ON (
                    T.SRH_UNIT = D.SRD_UNIT
                AND T.SRH_GROUP = D.SRD_GROUP
                AND T.SRH_CODE = D.SRD_CODE
                AND T.SRH_KEY_4 = D.SRD_KEY_4
                AND T.SRH_KEY_5 = D.SRD_KEY_5
                )
                WHEN MATCHED THEN
                UPDATE
                    SET 
                        T.SRH_LOT_NO = D.SRD_LOT_NO
                        , T.SRH_PROCESS = D.SRD_PROCESS
                        , T.SRH_MC = D.SRD_MC
                        , T.SRH_OPERATOR = D.SRD_OPERATOR
                        , T.SRH_TIME = D.SRD_TIME
                        , T.SRH_DATA_FILENAME = D.SRD_FILENAME
                WHEN NOT MATCHED THEN
                INSERT (
                    T.SRH_UNIT
                    , T.SRH_GROUP
                    , T.SRH_CODE
                    , T.SRH_DATE
                    , T.SRH_LOT_NO
                    , T.SRH_PROCESS
                    , T.SRH_MC
                    , T.SRH_KEY_1
                    , T.SRH_KEY_2
                    , T.SRH_KEY_3
                    , T.SRH_KEY_4
                    , T.SRH_KEY_5
                    , T.SRH_OPERATOR
                    , T.SRH_TIME
                    , T.SRH_DATA_FILENAME
                )
                VALUES (
                    D.SRD_UNIT
                    , D.SRD_GROUP
                    , D.SRD_CODE
                    , D.SRD_DATE
                    , D.SRD_LOT_NO
                    , D.SRD_PROCESS
                    , D.SRD_MC
                    , D.SRD_KEY_1
                    , D.SRD_KEY_2
                    , D.SRD_KEY_3
                    , D.SRD_KEY_4
                    , D.SRD_KEY_5
                    , D.SRD_OPERATOR
                    , D.SRD_TIME
                    , D.SRD_FILENAME
                )
            """
            query_list.append(header)

            seq = 1
            while seq <= 18:
                str_seq = str(seq)
                val = df.loc[i, str_seq]
                detail = f"""
                MERGE INTO SMF_RECORD_DETAIL T
                USING (
                    SELECT 
                            'CFM' AS SRD_UNIT
                            , '{smf_group}' AS SRD_GROUP
                            , '{code}' AS SRD_HEADER_CODE
                            , '-' AS SRD_KEY_1
                            , '-' AS SRD_KEY_2
                            , '-' AS SRD_KEY_3
                            , '{str_seq}' AS SRD_SEQ
                            , '1' AS SRD_ROUND
                            , '-' AS SRD_VALUE_CHR
                            , {val} AS SRD_VALUE_NUM
                    FROM DUAL
                ) D
                ON (
                    T.SRD_UNIT = D.SRD_UNIT
                AND T.SRD_GROUP = D.SRD_GROUP
                AND T.SRD_HEADER_CODE = D.SRD_HEADER_CODE
                AND T.SRD_SEQ = D.SRD_SEQ
                AND T.SRD_ROUND = D.SRD_ROUND
                AND T.SRD_VALUE_CHR = D.SRD_VALUE_CHR
                )
                WHEN MATCHED THEN
                UPDATE
                    SET T.SRD_VALUE_NUM = D.SRD_VALUE_NUM
                WHEN NOT MATCHED THEN
                INSERT (
                    T.SRD_UNIT
                    , T.SRD_GROUP
                    , T.SRD_HEADER_CODE
                    , T.SRD_KEY_1
                    , T.SRD_KEY_2
                    , T.SRD_KEY_3
                    , T.SRD_SEQ
                    , T.SRD_ROUND
                    , T.SRD_VALUE_CHR
                    , T.SRD_VALUE_NUM
                )
                VALUES (
                    D.SRD_UNIT
                    , D.SRD_GROUP
                    , D.SRD_HEADER_CODE
                    , D.SRD_KEY_1
                    , D.SRD_KEY_2
                    , D.SRD_KEY_3
                    , D.SRD_SEQ
                    , D.SRD_ROUND
                    , D.SRD_VALUE_CHR
                    , D.SRD_VALUE_NUM
                )
                """
                query_list.append(detail)
                seq += 1

            for query in query_list:
                str_error = self.config.save_irpt(sql_command=query)
                if str_error == "":
                    self.total_commit += 1
            
            # break

    def get_max_header(self, str_date, str_mc):
        sql_command = f"""
            SELECT MAX(T.SRH_CODE) AS HC
            FROM SMF_RECORD_HEADER T
            WHERE T.SRH_CODE LIKE '{str_date}%
                  AND T.SRH_MC = '{str_mc}'
        """
        hc_df = self.config.select_irpt(sql_command=sql_command)
        max_hc = hc_df.loc[0, 'HC']
        if max_hc == None:
            max_hc = str_date + '0001'
        else:
            max_hc = int(max_hc) + 1

        return max_hc

if __name__ == '__main__':
    app = SetRdes()
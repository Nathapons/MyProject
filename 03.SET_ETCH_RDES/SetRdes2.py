from settings import Settings
import os
import shutil
import datetime
import pandas as pd


class SetRdes(Settings):
    def __init__(self):
        self.total_commit = 0
        super().write_debug(word="Start Program")
        self.get_file_dir()
        super().write_debug(word="End Program")

    def get_file_dir(self):
        # Get filename in Folder IN
        file_dir = os.listdir(super().source_path)

        if len(file_dir) > 0:
            for file in file_dir:
                now = datetime.datetime.now()
                now_format = now.strftime('%Y%m%d_%H%m_')
                new_file = now_format + file
                file_in = os.path.join(super().source_path, file)
                file_out = os.path.join(super().out_path, new_file)
                file_error = os.path.join(super().error_path, new_file)

                try:
                    if str(file).endswith(super().file_type) and (not str(file).startswith('~$')):
                        super().write_debug(word=f"     Open file : {str(file)}")
                        self.des_program(file_path=file_in)

                        if os.path.isfile(file_out):
                            os.remove(file_out)
                        shutil.move(file_in, file_out)
                except Exception as e:
                    super().write_debug(word=f"     Error=> {e}")
                    if os.path.isfile(file_error):
                        os.remove(file_error)
                    shutil.move(file_in, file_error)
                finally:
                    super().write_debug(word=f"     Total commit: {self.total_commit} data")
                    super().write_debug(word=f"     End Read file : {str(file)}")
                    
        else:
            super().write_debug(word="     There is no FILE. Wait until next event")

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
                            date_form = date.strftime('%m/%d/%Y')
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
                        # super().write_debug(word=f"      Convert at row {i+10} has error {str_error}")
                        continue 

                    # Etching Upper
                    has_upper = True
                    seq = 4
                    col = 5
                    while col <= 9:
                        weight_before = df.iloc[i-3, col]
                        weight_after = df.iloc[i-2, col]
                        etch_upper = ""

                        if isinstance(weight_before, float) == False or isinstance(weight_after, float) == False:
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

                        if isinstance(weight_before, float) == False or isinstance(weight_after, float) == False:
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
                        chem = df.iloc[i, col]

                        if isinstance(chem, float) == False:
                            has_chem = False
                        else:
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

                        is_nan = False
                        for val in datarow.values():
                            if str(val) == 'nan':
                                is_nan = True

                        if is_nan == False:
                            datatable.append(datarow)

        rdes = pd.DataFrame(datatable)
        print(rdes)
        if len(rdes) > 0:
            smf_group = '0001'
            super().write_debug(word="     Start insert to RDES Datatabase")
            self.insert_smf_record(str_filename=filename, df=rdes, smf_group=smf_group)   
            super().write_debug(word="     End insert to RDES Datatabase")

    def insert_smf_record(self, str_filename, df, smf_group):
        query_list = []
        
        for i in range(len(df)):
            date = df.loc[i, 'date']
            machine = df.loc[i, 'machine']
            date_sql = df.loc[i, 'date_sql']
            factory = df.loc[i, 'factory']
            operator = df.loc[i, 'operator']
            time = df.loc[i, 'time']
            code = super().get_max_header(
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
                str_error = super().save_irpt(sql_command=query)
                if str_error == "":
                    self.total_commit += 1
                else:
                    print(query)
            query_list.clear()

    def get_max_header(self, str_date, str_mc):
        sql_command = f"""
            SELECT MAX(T.SRH_CODE) AS HC
            FROM SMF_RECORD_HEADER T
            WHERE T.SRH_CODE LIKE '{str_date}%
                  AND T.SRH_MC = '{str_mc}'
        """
        hc_df = super().select_irpt(sql_command=sql_command)
        max_hc = hc_df.loc[0, 'HC']
        if max_hc == None:
            max_hc = str_date + '0001'
        else:
            max_hc = int(max_hc) + 1

        return max_hc


if __name__ == '__main__':
    app = SetRdes()
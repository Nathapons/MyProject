from pandas.io import excel
import pandas as pd
import os
from shutil import copyfile
from settings import Settings


class Ort:
    def __init__(self):
        self.config = Settings()
        self.total_copy = 0
        self.error_copy = 0

    def program(self):
        self.config.write_debug(word="Start Program")
        try:
            self.get_ort_excel()
        except Exception as error:
            str_error = str(error)
            self.config.write_debug(word=f"  Program define error => {str_error}")

        self.config.write_debug(word=f"  Total copy file => {self.total_copy}")
        self.config.write_debug(word=f"  Error copy file => {self.error_copy}")
        self.config.write_debug(word="End Program")

    def get_ort_excel(self):
        ort_dt = []
        ort_status = self.config.ort_path
        
        for file in os.listdir(ort_status):
            if file.startswith('ORT') and file.upper().endswith('XLSB'):
                ort_excel = os.path.join(ort_status, file)
                sheetnames = self.get_sheetnames(excel_file=ort_excel)
                ort_dt = self.read_data(ort_dt, ort_excel, sheetnames)

        ort_df = pd.DataFrame(ort_dt, columns=['barcode', 'product', 'lotno', 'item_test'])
        if len(ort_df) > 0:
            self.get_filename(ort_df)

    def get_sheetnames(self, excel_file):
        sheetnames = []
        try:
            xl = pd.ExcelFile(excel_file)
            sheetnames = [name for name in xl.sheet_names if name not in self.config.not_read_list]
        except Exception as error:
            str_filename = os.path.basename(excel_file)
            self.config.write_debug(word=f'  Error for getting sheet list filename => {str_filename}')

        return sheetnames

    def read_data(self, ort_dt, excel_file, sheetnames): 
        engine = None
        if excel_file.upper().endswith('XLSB'):
            engine = 'pyxlsb'

        for sheetname in sheetnames:
            try:
                df = pd.read_excel(excel_file, sheet_name=sheetname, header=1, dtype=str, engine=engine)
                
                for i in df.index:
                    barcode = str(df['S/N'][i])
                    product_name = str(df['P/D name'][i]).strip()

                    if "-" not in product_name and "Z" not in product_name:
                        first_path = product_name[0:3]
                        second_path = product_name[3:]
                        full_product_name = str(first_path) + "-" + str(second_path)
                    elif "-" not in product_name:
                        first_path = product_name[0:4]
                        second_path = product_name[4:]
                        full_product_name = str(first_path) + "-" + str(second_path)
                    else:
                        full_product_name = product_name
                    
                    lotno = str(df['lot no. for OST'][i])
                    item_test = str(df['Item test'][i])

                    row = [barcode, full_product_name, lotno, item_test]
                    if 'nan' not in row:
                        ort_dt.append(row)

            except Exception as error:
                str_file = os.path.basename(excel_file)
                str_error = str(error)
                self.config.write_debug(word=f'  {str_file} has error to get dataframe=> {str_error} at {sheetname}')

        return ort_dt

    def get_filename(self, ort_df):
        self.config.write_debug(word='     Start Move file in R2-40-131 folder') 
        yamaha_path = self.config.yamaha
        yamaha_files = os.listdir(yamaha_path)
        for file in yamaha_files:
            if file.startswith('A') and file.upper().endswith('CSV'):
                barcode_no = file.split('_')[0]
                self.check_item_sort(ort_df, barcode_no, file, yamaha_path)
        self.config.write_debug(word='     End Move file in R2-40-131 folder') 

        self.config.write_debug(word='     Start Move file in W-40-112 folder') 
        nidech_path = self.config.nidech
        nidech_files = os.listdir(nidech_path)
        for file in nidech_files:
            if file.startswith('A') and file.upper().endswith('CSV'):
                barcode_no = file.split('_')[0]
                self.check_item_sort(ort_df, barcode_no, file, nidech_path)
        self.config.write_debug(word='     End Move file in W-40-112 folder')

    def check_item_sort(self, ort_df, barcode_no, file, machine_path):
        index_filter = 0
        is_found = False
        for i in range(len(ort_df)):
            barcode_ort = ort_df.loc[i, "barcode"]
            if barcode_no == barcode_ort:
                index_filter = i
                is_found = True

        if is_found:
            product = ort_df.loc[index_filter, "product"]
            lotno = ort_df.loc[index_filter, "lotno"]
            item_test = ort_df.loc[index_filter, "item_test"]
            source = os.path.join(machine_path, file)
            self.create_folder_and_copy_file(product, lotno, item_test, source)

    def create_folder_and_copy_file(self, product, lotno, item_test, source):
        product_path = os.path.join(self.config.destination, product)
        item_test_path = os.path.join(product_path, item_test)
        lotno_path = os.path.join(item_test_path, lotno)
        is_product_path = os.path.isdir(product_path)
        is_item_test_path = os.path.isdir(item_test_path)
        is_lotno_path = os.path.isdir(lotno_path)

        try:
            if is_product_path == False:
                os.mkdir(product_path)
            if is_item_test_path == False:
                os.mkdir(item_test_path)
            if is_lotno_path == False:
                os.mkdir(lotno_path)
        except Exception as error:
            str_error = str(error)
            self.config.write_debug(f'      Server Access is denied when create folder => {str_error}')

        filename = os.path.basename(source)
        destination = os.path.join(lotno_path, filename)
        is_file_exist = os.path.isfile(destination)
        
        if is_file_exist == False:
            try:
                copyfile(source, destination)
                self.total_copy += 1
            except Exception as error:
                str_error = str(error)
                self.error_copy += 1
                self.config.write_debug(f'     Copy file Error => {str_error}')


if __name__ == '__main__':
    app = Ort()
    app.program()

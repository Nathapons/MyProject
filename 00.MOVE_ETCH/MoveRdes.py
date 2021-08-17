import openpyxl as xl
from settings import Settings
import os
import shutil
from datetime import datetime


class MoveRdes:
    def __init__(self):
        # Call class from another file
        self.setting = Settings()
        self.oracle, self.config = self.setting.read_config()
 
        try:
            # Connect Oracle
            self.setting.write_debug(word="Start Program", log_path=self.config['log'])
            
            # Start Program
            self.get_in_folder()

        except Exception as error:
            self.setting.write_debug(word=str(error), log_path=self.config['log'])
        finally:
            self.setting.write_debug(word="End Program", log_path=self.config['log'])

    def get_in_folder(self):
        now = datetime.now()
        now_format = now.strftime('%Y%m')
        source = self.config['source']
        destination = self.config['destination']
        file_type = self.config['file_type']

        for folder1 in os.listdir(source):
            folder1_path = os.path.join(source, folder1)
            if 'MASTER' not in folder1.upper():

                for folder2 in os.listdir(folder1_path):
                    folder2_path = os.path.join(folder1_path, folder2)
                    if os.path.isdir(folder2_path):

                        for file in os.listdir(folder2_path):
                            if file.endswith(file_type) and not file.startswith('~$') and 'MASTER' not in file.upper():
                                file_format = folder2 + file.split('.')[0]
                                
                                new_file = folder1 + "_" + file
                                server_file = os.path.join(destination, new_file)
                                source_file = os.path.join(folder2_path, file)
                                complete_folder = os.path.join(folder2_path, 'COMPLETE')
                                complete_file = os.path.join(folder2_path, 'COMPLETE', file)

                                # Copy to Server
                                if os.path.isfile(server_file):
                                    os.remove(server_file)
                                shutil.copy2(src=source_file, dst=server_file)
                                
                                # Copy to Complete
                                try:
                                    now_format = int(datetime.now().strftime('%Y%m'))
                                    str_file_format = str(folder2) + str(file.split('.')[0])
                                    file_format = int(str_file_format)
                                    # print(now_format, file_format)
                                    if os.path.isdir(complete_folder) == False:
                                        os.mkdir(complete_folder)

                                    if os.path.isfile(complete_file):
                                        os.remove(complete_file)

                                    if file_format < int(now_format):
                                        shutil.move(src=source_file, dst=complete_file)
                                    elif file_format == int(now_format):
                                        shutil.copy2(src=source_file, dst=complete_file)

                                except Exception as error:
                                    str_error = str(error)
                                    self.setting.write_debug(
                                        word=f" Get folder error => {str_error}", 
                                        log_path=self.config['log']
                                    )


if __name__ == '__main__':
    app = MoveRdes()
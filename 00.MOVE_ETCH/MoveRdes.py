import openpyxl as xl
from settings import Settings
import os
import shutil
from datetime import datetime


class MoveRdes:
    def __init__(self):
        # Call class from another file
        self.settings = Settings
 
        try:
            # Connect Oracle
            self.oracle, self.config = self.settings.read_config(self)
            self.settings.write_debug(self, word="Start Program", log_path=self.config['log'])
            
            # Start Program
            self.get_in_folder()

        except Exception as error:
            self.settings.write_debug(self, word=str(error), log_path=self.config['log'])
        finally:
            self.settings.write_debug(self, word="End Program", log_path=self.config['log'])

    def get_in_folder(self):
        source = self.config['source']
        destination = self.config['destination']
        file_type = self.config['file_type']

        for folder1 in os.listdir(source):
            folder1_path = os.path.join(source, folder1)
            if 'MASTER' not in folder1.upper() and os.path.isdir(folder1_path):

                for folder2 in os.listdir(folder1_path):
                    folder2_path = os.path.join(folder1_path, folder2)
                    if os.path.isdir(folder2_path):

                        for file in os.listdir(folder2_path):
                            if file.endswith(file_type) and not file.startswith('~$') and 'MASTER' not in file.upper():
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
                                month_now = int(datetime.now().strftime('%m'))
                                file_monthnum = int(file.split('.')[0])
                                if os.path.isdir(complete_folder) == False:
                                    os.mkdir(complete_folder)

                                if os.path.isfile(complete_file):
                                    os.remove(complete_file)

                                if file_monthnum < month_now:
                                    shutil.move(src=source_file, dst=complete_file)
                                elif file_monthnum == month_now:
                                    shutil.copy2(src=source_file, dst=complete_file)


if __name__ == '__main__':
    app = MoveRdes()
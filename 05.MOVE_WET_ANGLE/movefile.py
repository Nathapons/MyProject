from settings import Settings
import os
import shutil


class MoveFile(Settings):
    def __init__(self):
        self.total_move = 0
        try:
            # super is inheritance from Settings class of settings.py
            super().__init__()
            super().write_debug(word="Start Program")
            self.get_path()

        except Exception as error:
            str_error = str(error)
            super().write_debug(word=str_error)
        finally:
            super().write_debug(word="End Program")
            super().write_debug(word="Total move file: {}".format(self.total_move))

    def get_path(self):
        file_dict = super().get_file()
        src = file_dict.get('src') 
        
        for src_file in os.listdir(src):
            self.move_copy_file(src_file=src_file)

    def move_copy_file(self, src_file):
        file_dict = super().get_file()
        src = file_dict.get('src')
        des = file_dict.get('des')
        file_type = file_dict.get('type')
        src_file_path = os.path.join(src, src_file)

        if src_file not in os.listdir(des) and src_file.endswith(file_type) and os.path.isfile(src_file_path):
            # Copy to IOT Computer
            err_copy = ""
            try:
                src_path = os.path.join(src, src_file)
                dst_path = os.path.join(des, src_file)
                shutil.copy2(src=src_path, dst=dst_path)
            except Exception as e:
                err_copy = str(e)
                super().write_debug(word="  Error copy to IOT => {}".format(src_file))

            # Move file to compelete
            err_move = ""
            try:
                src_path = os.path.join(src, src_file)
                complete_path = os.path.join(src, "COMPLETE")
                dst_path = os.path.join(src, "COMPLETE", src_file)

                if os.path.isdir(complete_path) == False:
                    os.mkdir(complete_path)
                if os.path.isfile(dst_path):
                    os.remove(dst_path)
                shutil.move(src=src_path, dst=dst_path)
            except Exception as e:
                err_move = str(e)
                super().write_debug(word="  Move file error => {}".format(src_file))
            
            if err_copy != "" and err_move != "":
                self.total_move += 1
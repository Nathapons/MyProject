from settings import Settings
from os import path, remove, listdir
from shutil import move
from datetime import datetime


class SetDb(Settings):
    def __init__(self):
        super().__init__()
        self.total_commit = 0
        super().write_debug(word="Start program")
        self.get_files()
        super().write_debug(word="Total commit: {}".format(self.total_commit))
        super().write_debug(word="End program")

    def get_files(self):
        file = super().get_file()
        ins = file.get('in')
        file_dirs = listdir(ins)

        if len(file_dirs) > 0:
            exit

        for file in file_dirs:
            now = datetime.now()
            now_format = now.strftime('%Y%m%d_%H%m_')
            new_file = now_format + file
            file_in = path.join(self.config.source_path, file)
            file_out = path.join(self.config.out_path, new_file)
            file_error = path.join(self.config.error_path, new_file)

            try:
                super().write_debug(word="     Start Read file: {}".format(file))
            except Exception as err:
                str_err = str(err)
                super().write_debug(word=f"     Error=> {str_err}")
            finally:
                super().write_debug(word="     End Read file")
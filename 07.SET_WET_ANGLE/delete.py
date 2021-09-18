from settings import Settings
import os


class DeleteFile(Settings):
    def __init__(self):
        super().__init__()

    def get_path(self):
        file_dict = super().get_file()
        out = file_dict.get('out')
        error = file_dict.get('error')
        log = file_dict.get('log')
        file_type = file_dict.get('type')

        self.delete_file(out=out, error=error, type=file_type)
        self.delete_file(log=log, type='.txt')

    def delete_file(self, type, **path):
        for source in path:
            path_name = path[source]
            for file in os.listdir(path_name):
                file_path = os.path.join(path_name, file)
                if os.path.isfile(file_path) and file.endswith(type):
                    os.remove(file_path)


if __name__ == '__main__':
    obj = DeleteFile()
    obj.get_path()

    
from ost import Ost 


if __name__ == '__main__':
    obj = Ost()
    src, des, type, log_path = obj.get_json()
    obj.copy_file(source=src, destination=des, file_type=type)
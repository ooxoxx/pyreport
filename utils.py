import pyodbc
# def singleton(cls):
#     _instance = {}

#     def inner():
#         if cls not in _instance:
#             _instance[cls] = cls()
#         return _instance[cls]

#     return inner


# @singleton # 这里考虑不需要单例模式
class Cursor:
    def __init__(self, conn_str):
        self.conn_str = conn_str

    def get_cursor(self):
        try:
            cnxn = pyodbc.connect(self.conn_str)
            self._cursor = cnxn.cursor()
        except Exception:
            raise Exception('数据库访问不到')
        return self._cursor


def ms_access(mdb_filepath):
    conn_str = (
        'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={mdb_filepath};' 
        )

    return conn_str


if __name__ == '__main__':
    print(Cursor(ms_access('data/ClouMeterData_original.mdb')).get_cursor())
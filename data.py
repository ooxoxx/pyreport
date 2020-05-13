import pyodbc
import numpy as np
import textwrap
import configparser as cp


class _AccessCursor(object):

    def __init__(self, mdb_filepath):
        self.conn_str = (
            'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={mdb_filepath};'
            )

    def get_cursor(self):
        try:
            cnxn = pyodbc.connect(self.conn_str)
            cursor = cnxn.cursor()
        except Exception:
            raise Exception('database unconnected.')
        return cursor


class _Data(object):

    def __new__(cls, *args, **kargs):
        config = cp.ConfigParser()
        config.read('./config.ini')
        mdb_filepath = config.get('access', 'mdb_filepath')
        cls._cursor = _AccessCursor(mdb_filepath).get_cursor()
        instance = object.__new__(cls)
        return instance

    def __init__(self, id_data):
        self._meter_id = id_data.meter_id

    def read(self):
        raise Exception("Data.read() not implemented.")


class IdData(_Data):

    def __init__(self, meter_address):
        self._meter_address = meter_address
        sql = textwrap.dedent(f"""
            SELECT PK_LNG_METER_ID
            FROM METER_INFO
            WHERE AVR_ADDRESS = '{self._meter_address}'
            """)
        data = self._cursor.execute(sql).fetchone()
        self._meter_id = data[0]

    def read(self):
        return np.array([self._meter_address]).reshape(1, 1)

    @property
    def meter_id(self):
        return self._meter_id


class DeviationData(_Data):
    def __init__(self, id_data, power_type, component, error_type='0'):
        self._meter_id = id_data.meter_id
        self._error_type = error_type
        power_type_dict = {
                'active': '0',
                'reversed': '1',
                'reactive': '2'
                }
        self._power_type = power_type_dict[power_type]
        component_dict = {
                'unbalanced': '<>',
                'balanced': "=",
                }
        self._component = component_dict[component]

    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_ERROR_MORE, AVR_PROJECT_NO
                FROM METER_ERROR
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND CHR_ERROR_TYPE = '{self._error_type}'
                AND CHR_POWER_TYPE = '{self._power_type}'
                AND CHR_COMPONENT {self._component} '1'
                ORDER BY AVR_PROJECT_NO
                """)
        data = self._cursor.execute(sql).fetchall()
        deviation = np.array([e[0].split('|') for e in data])
        self._project_no = np.array([e[1] for e in data])
        return deviation


class JiduData(_Data):

    def __init__(self, id_data):
        self._meter_id = id_data.meter_id

    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_VALUE
                FROM METER_COMMUNICATION
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_PROJECT_NO >= '00500'
                AND AVR_PROJECT_NO <= '00504'
                ORDER BY AVR_PROJECT_NO
                """)
        data = self._cursor.execute(sql).fetchall()
        data = np.array([e[0].split('|') for e in data])
        data = data[[3, 2, 1, 4, 0]]
        data = data[:, :3].astype('float').T
        尖峰平谷 = data[:, :4]
        分时 = 尖峰平谷.sum(axis=1)
        总 = data[:, -1]
        误差 = 总 - 分时
        temp1 = np.c_[尖峰平谷, 分时, 总]
        temp = [f'{e:.2f}' for e in temp1.ravel()]
        temp = np.array(temp).reshape(temp1.shape)
        误差 = np.array(list(map(lambda e: f'{e:+.2f}', 误差)))
        data = np.c_[temp, 误差]
        return data


class XuliangData(_Data):

    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_VALUE
                FROM METER_COMMUNICATION
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_PROJECT_NO LIKE '01_11'
                ORDER BY AVR_PROJECT_NO DESC
                """)
        data = self._cursor.execute(sql).fetchall()
        temp = np.array([e[0].split('|') for e in data])
        标准 = temp[:, 0]
        实际 = np.array(list(map(lambda e: f'{float(e):.4f}', temp[:, 1])))
        误差 = np.array(list(map(lambda e: f'{float(e):+.2f}', temp[:, 2])))
        return np.c_[标准, 实际, 误差]


<<<<<<< HEAD
if __name__ == '__main__':
    print(DeviationData(meter_id='6718967118237250304', 
        power_type='active', component='balanced').read())
    #print(IdData('910003622190').read())
=======
class BianchaData(_Data):

    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_DATAS_1, AVR_DATAS_2
                FROM METER_CONSISTENCY_DATA
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_ITEM_TYPE LIKE '2_%'
                ORDER BY AVR_ITEM_TYPE
                """)
        data = self._cursor.execute(sql).fetchall()
        data = np.array(data)
        left = data[:, 0]
        right = data[:, 1]
        left = np.array(list(map(lambda x: x.split('|')[:3], left)))
        right = np.array(list(map(lambda x: x.split('|')[:3], right)))
        temp = right[:, -1].astype('float') - left[:, -1].astype('float')
        误差变差 = np.array(list(map(lambda x: f"{x:+.4f}", temp)))
        修约后 = np.array(list(map(lambda x: f"{x:+.2f}", temp)))
        return np.c_[left, right, 误差变差, 修约后]
>>>>>>> 75ab18d810fcdbfea01c8eabe3ce7a631ed777f1


class YizhixingData(_Data):

    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_DATAS_1
                FROM METER_CONSISTENCY_DATA
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_ITEM_TYPE LIKE '1_%'
                ORDER BY AVR_ITEM_TYPE """)
        data = self._cursor.execute(sql).fetchall()
        data = np.array(list(map(lambda x: x[0].split('|')[:3], data)))
        self.avr = data.astype('float').mean(axis=1)
        return data


class YizhixingMeanData(_Data):

    def __init__(self, yzx_list):
        avr_list = list(map(lambda x: x.avr, yzx_list))
        self._avr_matrix = np.array(avr_list).T
        # print('avr=', self._avr_matrix)

    def read(self):
        同批次平均 = self._avr_matrix.mean(axis=1)
        # print('tpc=', 同批次平均.reshape(3,1))
        变化值_matrix = self._avr_matrix - 同批次平均.reshape(3, 1)
        # print('bhz_matrix=', 变化值_matrix)
        temp_变化值 = np.r_[tuple(变化值_matrix[:, i] for i in range(3))]
        # print(temp_变化值)
        temp_同批次 = np.r_[(同批次平均,)*3]
        col1 = list(map(lambda x: f"{x:+.4f}", temp_变化值))
        col2 = list(map(lambda x: f"{x:+.2f}", temp_变化值))
        col3 = list(map(lambda x: f"{x:+.4f}", temp_同批次))
        return np.c_[col1, col2, col3]


class FuzaidianliuData(_Data):

    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_DATAS_1, AVR_DATAS_2
                FROM METER_CONSISTENCY_DATA
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_ITEM_TYPE LIKE '3_%'
                ORDER BY AVR_ITEM_TYPE
                """)
        data = self._cursor.execute(sql).fetchall()
        data = np.array(data)
        left = data[:, 0]
        right = data[:, 1]
        left = np.array(list(map(lambda x: x.split('|')[:3], left)))
        right = np.array(list(map(lambda x: x.split('|')[:3], right)))
        self.left = left
        self.right = right
        return np.r_[left, right[::-1]]


class FuzaidianliuAggrData(_Data):

    def __init__(self, fzdl_data):
        self._left = fzdl_data.left[:, -1].astype('float')
        self._right = fzdl_data.right[:, -1].astype('float')

    def read(self):
        diff = self._right - self._left
        col1 = list(map(lambda x: f"{x:+.4f}", diff))
        col2 = list(map(lambda x: f"{x:+.2f}", diff))
        res = np.c_[col1, col2]
        return res


if __name__ == '__main__':
    id_data = IdData('910003622190')
    # print(DeviationData(id_data, 'active', 'balanced').read())
    # print(JiduData(id_data).read())
    # print(XuliangData(id_data).read())
    # print(BianchaData(id_data).read())
    # print(IdData('910003622190').read())
    # yd = YizhixingData(id_data)
    # print(yd.read())
    # lst = []

    # class Test:
        # pass
    # for i in range(3):
        # e = Test()
        # e.avr = np.array([0.2, 0.3, 0.4])*(i+1)
        # lst.append(e)
    # print(YizhixingMeanData(lst).read())
    data = FuzaidianliuData(id_data)
    print(data.read())
    print(FuzaidianliuAggrData(data).read())

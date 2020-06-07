import numpy as np
import textwrap
import configparser as cp
import platform
from writer_singleton import writer


class _WindowsAccess(object):
    def __init__(self, mdb_filepath):
        driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
        self.conn_str = (f'DRIVER={driver};DBQ={mdb_filepath};')

    def get_cursor(self):
        import pyodbc
        try:
            cnxn = pyodbc.connect(self.conn_str)
            cursor = cnxn.cursor()
        except Exception:
            raise Exception('database unconnected.')
        return cursor


class _LinuxAccess(object):
    def __init__(self, mdb_filepath):
        self._mdb_filepath = mdb_filepath

    def get_cursor(self):
        import jaydebeapi
        ucanaccess_jars = [
            "lib/ucanaccess-5.0.0.jar",
            "lib/lib/hsqldb-2.5.0.jar",
            "lib/lib/jackcess-3.0.1.jar",
            "lib/lib/commons-lang3-3.8.1.jar",
            "lib/lib/commons-logging-1.2.jar",
        ]
        classpath = ":".join(ucanaccess_jars)
        cncx = jaydebeapi.connect("net.ucanaccess.jdbc.UcanaccessDriver",
                                  f"jdbc:ucanaccess://{self._mdb_filepath}",
                                  None, classpath)
        cursor = cncx.cursor()
        return cursor


def initialize_cursor():
    config = cp.ConfigParser()
    config.read('./config.ini')
    mdb_filepath = config.get('access', 'mdb_filepath')
    # multi platform support
    osname = platform.system()
    if osname == "Linux":
        access = _LinuxAccess
    elif osname == "Windows":
        access = _WindowsAccess
    else:
        raise Exception('your os not supported.')
    return access(mdb_filepath).get_cursor()


_cursor = initialize_cursor()

set_template = writer.set_template


def save():
    writer.save()


class Data(object):
    def __init__(self, id_data):
        self._meter_id = id_data.meter_id
        self._content = None

    def read(self) -> np.ndarray:
        raise NotImplementedError

    def fill(self, tab_pos, start_cell):
        self._content = self.read()
        writer.write(tab_pos, self.content, start_cell)

    @property
    def content(self):
        if self._content is None:
            print('WARNING: report result before original record!')
            return self.read()
        else:
            return self._content


class IdData(Data):
    def __init__(self, meter_address):
        self._meter_address = meter_address
        self._meter_id = None
        self._meter_class = None

    def read(self):
        return np.array([self._meter_address]).reshape(1, 1)

    @property
    def meter_id(self):
        if self._meter_id is None:
            self.query_info()
        return self._meter_id

    @property
    def meter_class(self):
        if self._meter_class is None:
            self.query_info()
        return self._meter_class

    def query_info(self):
        sql = textwrap.dedent(f"""
            SELECT PK_LNG_METER_ID, AVR_AR_CLASS
            FROM METER_INFO
            WHERE AVR_ADDRESS = '{self._meter_address}'
            """)
        _cursor.execute(sql)
        data = _cursor.fetchone()
        self._meter_id = data[0][0]
        self._meter_class = data[0][1]


class DeviationData(Data):
    def __init__(self, id_data, power_type, component, error_type='0'):
        self._meter_id = id_data.meter_id
        self._error_type = error_type
        power_type_dict = {'active': '0', 'reversed': '1', 'reactive': '2'}
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
        _cursor.execute(sql)
        data = _cursor.fetchall()
        deviation = np.array([e[0].split('|') for e in data])
        self._project_no = np.array([e[1] for e in data])
        return deviation


class JiduData(Data):
    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_VALUE
                FROM METER_COMMUNICATION
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_PROJECT_NO >= '00500'
                AND AVR_PROJECT_NO <= '00504'
                ORDER BY AVR_PROJECT_NO
                """)
        _cursor.execute(sql)
        data = _cursor.fetchall()
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


class XuliangData(Data):
    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_VALUE
                FROM METER_COMMUNICATION
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_PROJECT_NO LIKE '01_11'
                ORDER BY AVR_PROJECT_NO DESC
                """)
        _cursor.execute(sql)
        data = _cursor.fetchall()
        temp = np.array([e[0].split('|') for e in data])
        标准 = temp[:, 0]
        实际 = np.array(list(map(lambda e: f'{float(e):.4f}', temp[:, 1])))
        误差 = np.array(list(map(lambda e: f'{float(e):+.2f}', temp[:, 2])))
        return np.c_[标准, 实际, 误差]


class BianchaData(Data):
    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_DATAS_1, AVR_DATAS_2
                FROM METER_CONSISTENCY_DATA
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_ITEM_TYPE LIKE '2_%'
                ORDER BY AVR_ITEM_TYPE
                """)
        _cursor.execute(sql)
        data = _cursor.fetchall()
        data = np.array(data)
        left = data[:, 0]
        right = data[:, 1]
        left = np.array(list(map(lambda x: x.split('|')[:3], left)))
        right = np.array(list(map(lambda x: x.split('|')[:3], right)))
        temp = right[:, -1].astype('float') - left[:, -1].astype('float')
        误差变差 = np.array(list(map(lambda x: f"{x:+.4f}", temp)))
        修约后 = np.array(list(map(lambda x: f"{x:+.2f}", temp)))
        return np.c_[left, right, 误差变差, 修约后]


class YizhixingData(Data):
    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_DATAS_1
                FROM METER_CONSISTENCY_DATA
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_ITEM_TYPE LIKE '1_%'
                ORDER BY AVR_ITEM_TYPE """)
        _cursor.execute(sql)
        data = _cursor.fetchall()
        data = np.array(list(map(lambda x: x[0].split('|')[:3], data)))
        self.avr = data.astype('float').mean(axis=1)
        return data


class YizhixingMeanData(Data):
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
        temp_同批次 = np.r_[(同批次平均, ) * 3]
        col1 = list(map(lambda x: f"{x:+.4f}", temp_变化值))
        col2 = list(map(lambda x: f"{x:+.2f}", temp_变化值))
        col3 = list(map(lambda x: f"{x:+.4f}", temp_同批次))
        return np.c_[col1, col2, col3]


class FuzaidianliuData(Data):
    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_DATAS_1, AVR_DATAS_2
                FROM METER_CONSISTENCY_DATA
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND AVR_ITEM_TYPE LIKE '3_%'
                ORDER BY AVR_ITEM_TYPE
                """)
        _cursor.execute(sql)
        data = _cursor.fetchall()
        data = np.array(data)
        left = data[:, 0]
        right = data[:, 1]
        left = np.array(list(map(lambda x: x.split('|')[:3], left)))
        right = np.array(list(map(lambda x: x.split('|')[:3], right)))
        self.left = left
        self.right = right
        return np.r_[left, right[::-1]]


class FuzaidianliuAggrData(Data):
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

import pyodbc
import numpy as np
import textwrap


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
    _cursor = _AccessCursor(r'data/ClouMeterData_original.mdb').get_cursor()

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
        data = np.c_[尖峰平谷, 分时, 总, 误差]
        temp = [f'{e:.2f}' for e in data.ravel()]
        data = np.array(temp).reshape(data.shape)
        return data


class InfoData(_Data):
    def __init__(self, meter_id, info):
        self._meter_id = meter_id
        self._info = info

    def read(self):
        pass


if __name__ == '__main__':
    id_data = IdData('910003622190')
    print(DeviationData(id_data, 'active', 'balanced').read())
    print(JiduData(id_data).read())
    # print(IdData('910003622190').read())

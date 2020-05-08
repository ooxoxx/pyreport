import pyodbc
import numpy as np
import textwrap


class AccessCursor:
    def __init__(self, mdb_filepath):
        self.conn_str = (
        'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={mdb_filepath};' 
        )

    def get_cursor(self):
        try:
            cnxn = pyodbc.connect(self.conn_str)
            _cursor = cnxn.cursor()
        except Exception:
            raise Exception('database unconnected.')
        return _cursor


class Data(object):
    _cursor = AccessCursor(r'data/ClouMeterData_original.mdb').get_cursor()

    def read(self):
        raise Exception("Data.read() not implemented.")


class IdData(Data):
    """docstring for MeterIDReader"""
    def __init__(self, meter_address):
        self._meter_address = meter_address

    def read(self):
        sql = textwrap.dedent(f"""
            SELECT PK_LNG_METER_ID
            FROM METER_INFO
            WHERE AVR_ADDRESS = '{self._meter_address}'
            """)
        data = self._cursor.execute(sql).fetchone()
        return data[0]

class DeviationData(Data):
    def __init__(self, meter_id, power_type, component, error_type='0'):
        self._meter_id = meter_id
        self._error_type = error_type
        power_type_dict = {
                'active': '0',
                'reactive': '1'
                }
        self._power_type = power_type_dict[power_type]
        component_dict = {
                'unbalanced': '',
                'balanced': "AND CHR_COMPONENT = '1'",
                }
        self.component_sql = component_dict[component]

    def read(self):
        sql = textwrap.dedent(f"""
                SELECT AVR_ERROR_MORE, AVR_PROJECT_NO
                FROM METER_ERROR
                WHERE FK_LNG_METER_ID = '{self._meter_id}'
                AND CHR_ERROR_TYPE = '{self._error_type}'
                AND CHR_POWER_TYPE = '{self._power_type}'
                {self.component_sql}
                ORDER BY AVR_PROJECT_NO
                """)
        data = self._cursor.execute(sql).fetchall()
        # [('+0.0539|+0.0568|+0.0553|+0.05', ), ('+0.0596|+0.0601|+0.0599|+0.05', ),...
        deviation = np.array([e[0].split('|') for e in data])
        project_no = np.array([e[1] for e in data]) 
        return deviation, project_no

class InfoData(Data):
    def __init__(self, meter_id, info):
        self._meter_id = meter_id
        self._info = info

    def read(self):
        pass


if __name__ == '__main__':
    print(DeviationData(meter_id='6718967118237250304', power_type='active', component='balanced').read())
    #print(IdData('910003622190').read())


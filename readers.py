from utils import Cursor, ms_access
import numpy as np
import textwrap


_cursor = Cursor(ms_access(r'data/ClouMeterData_original.mdb')).get_cursor()

class DeviationReader:
    def __init__(self, meter_id, power_type, component, error_type='0'):
        self._meter_id = meter_id
        self._error_type = error_type
        self._power_type = power_type
        self._component = component

    def read(self):
        component_sql = '' if self._component == None else f"AND CHR_COMPONENT = '{self._component}'"
        sql = textwrap.dedent(f"""
        SELECT AVR_ERROR_MORE, AVR_PROJECT_NO
        FROM METER_ERROR 
        WHERE FK_LNG_METER_ID = '{self._meter_id}'
        AND CHR_ERROR_TYPE = '{self._error_type}'
        AND CHR_POWER_TYPE = '{self._power_type}'
        {component_sql}
        ORDER BY AVR_PROJECT_NO
        """)
        data = _cursor.execute(sql).fetchall()
        # [('+0.0539|+0.0568|+0.0553|+0.05', ), ('+0.0596|+0.0601|+0.0599|+0.05', ),...
        deviation = np.array([e[0].split('|') for e in data])
        project_no = np.array([e[1] for e in data])
        return deviation, project_no



if __name__ == '__main__':
    print(DeviationReader(meter_id='6718967118237250304', power_type='0', component='1').read())
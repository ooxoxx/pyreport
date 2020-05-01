import pyodbc
import textwrap
import win32com.client as wc
import os, time


class AwesomeWriter:
    """
    Writer类
    """
    def __init__(self, file_path):
        self.file_path = os.path.normpath(os.path.join(os.getcwd(), file_path))
        self.wordApp = wc.gencache.EnsureDispatch('Word.Application')
        # self.wordApp.Visible = True
        try:
            self.document = self.wordApp.Documents.Open(self.file_path)
        except Exception:
            raise Exception('找不到报告模板嗷~ 放在./data/下面嗷~')

    def write_block(self, table_id, datal, start_pos, block_size):
        start_row, start_col = start_pos
        rsize, csize = block_size
        n = len(datal)
        assert rsize * csize == n

        table = self.document.Tables(table_id)
        for i in range(n):
            r = i // csize + start_row
            c = i % csize + start_col
            table.Cell(r, c).Range.Text = datal[i]

    def write_addr(self, table_id, meter_addr):
        table = self.document.Tables(table_id)
        table.Cell(1, 2).Range.Text = str(meter_addr)

    def save(self, target_path=None):
        if target_path is None:
            time_stamp = time.strftime('%m%d%H%M%S', time.localtime())
            target_path = self.file_path.split('.')[0] + time_stamp + '.docx'
        self.document.SaveAs(target_path)

    def __del__(self):
        if self.wordApp.Visible == False:
            self.wordApp.Quit(wc.constants.wdDoNotSaveChanges)

class AwesomeReader:
    """
    Reader类
    """
    def __init__(self, mdb_filepath):

        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={mdb_filepath};' 
            )
        try:
            cnxn = pyodbc.connect(conn_str)
            self.crsr = cnxn.cursor()
        except Exception:
            raise Exception('找不到数据库文件嗷~')

    def get_error(self, power_type=None, component=None,
        factor=None, I=None):
        sql = textwrap.dedent("""
            SELECT AVR_ERROR_MORE
            FROM METER_ERROR 
            WHERE FK_LNG_METER_ID = '{}'
            AND CHR_ERROR_TYPE = '0'
            AND CHR_POWER_TYPE = '{}'
            AND CHR_COMPONENT = '{}'
            AND AVR_POWER_FACTOR = '{}'
            AND AVR_IB_MULTIPLE = '{}'
            """).format(self.meter_id, power_type, component, factor, I)
        data = self.crsr.execute(sql).fetchall()
        if not data:
            res = ['']*4
        else:
            assert len(data) == 1
            res = data[0][0].split('|')
            assert len(res) == 4
        return res

    def set_id(self, meter_addr):
        sql = textwrap.dedent("""
            SELECT PK_LNG_METER_ID
            FROM METER_INFO 
            WHERE AVR_ADDRESS = '{}'
            """).format(meter_addr)
        data = self.crsr.execute(sql).fetchall()
        if len(data) == 0:
            raise Exception('数据库里没有这个电能表编号嗷~')
        self.meter_id = data[0][0]

    def check_meter_type(self, meter_addr):
        # 1.获取电能表等级和电流
        sql = textwrap.dedent("""
            SELECT AVR_AR_CLASS, AVR_IB
            FROM METER_INFO 
            WHERE AVR_ADDRESS = '{}'
            """).format(meter_addr)
        data = self.crsr.execute(sql).fetchall() # data = [(AVR_AR_CLASS, AVR_IB)]
        meter_class = data[0][0].strip()
        meter_ib = data[0][1].strip()

        # 2.判断电能表类型
        if meter_ib == '5(60)':
            meter_type = 2
        else:
            class_list = {'0.5S(2)':1, '1(2)':0, '0.2s(2)':3} #用字典来指定类型
            meter_type = class_list[meter_class]

        return meter_type


if __name__ == '__main__':
    ar = AwesomeReader(r'data/ClouMeterData_original.mdb')

    print(ar.check_meter_type('970000041003'))
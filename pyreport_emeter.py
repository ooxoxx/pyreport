from prutils import AwesomeReader, AwesomeWriter
import os


class AwesomeReport:
    """
    主类
    """
    def __init__(self, mdb_filepath, file_path, meter_list):
        self.ar = AwesomeReader(mdb_filepath)
        self.aw = AwesomeWriter(file_path)
        self.meter_list = meter_list
        # meter_type = self.ar.check_meter_type(meter_list[0])
        # file_path_list = [
        #     '1抽检原始记录A1级三相外置.docx', 
        #     '1抽检原始记录A0.5s级三相外置.docx', 
        #     '1抽检原始记录A1级三相内置.docx', 
        #     '1抽检原始记录A0.2s级三相外置.docx'
        #     ]
        # file_path = os.path.join(file_path, file_path_list[meter_type])

    def report_all(self):
        for i, meter_addr in enumerate(meter_list):
            self.report_meter(meter_addr, i)
            print(meter_addr, 'done')

    def report_meter(self, meter_addr, table_pos=0):
        self.ar.set_id(meter_addr)

        # 有功基本误差(总)
        block1 = []
        factor_list = ('1.0', '0.5L', '0.8C')
        I_list1 = ('Imax', '1.0Ib', '0.5Ib', '0.1Ib', '0.05Ib')
        I_list2 = ('Imax', '1.0Ib', '0.5Ib', '0.2Ib', '0.1Ib')
        for f in factor_list:
            I_list = I_list1 if f == '1.0' else I_list2
            for I in I_list:
                print(f,)
                block1 += self.ar.get_error(power_type='0',
                    component='1', factor=f, I=I)
        self.aw.write_block(4+2*table_pos, block1, start_pos=(5, 4), block_size=(15, 4))
        self.aw.write_addr(4+2*table_pos, meter_addr)

        # 有功基本误差(ABC相)
        block2 = []
        component_list = ['2', '3', '4']
        factor_list = ['1.0', '0.5L']
        I_list1 = ['Imax', '1.0Ib', '0.5Ib', '0.1Ib']
        I_list2 = ['Imax', '1.0Ib', '0.5Ib', '0.2Ib']
        for c in component_list:
            for f in factor_list:
                I_list = I_list1 if f == '1.0' else I_list2
                for I in I_list2:
                    block2 += self.ar.get_error(power_type='0',
                        component=c, factor=f, I=I)
        self.aw.write_block(4+2*table_pos, block2, start_pos=(23, 5), block_size=(24, 4))
        self.aw.write_addr(4+2*table_pos, meter_addr)

        # 反向有功基本误差(总)
        block3 = []
        factor_list = ('1.0', '0.5L', '0.8C')
        I_list1 = ('Imax', '1.0Ib', '0.5Ib', '0.1Ib', '0.05Ib')
        I_list2 = ('Imax', '1.0Ib', '0.5Ib', '0.2Ib', '0.1Ib')
        for f in factor_list:
            I_list = I_list1 if f == '1.0' else I_list2
            for I in I_list:
                block3 += self.ar.get_error(power_type='1',
                    component='1', factor=f, I=I)
        self.aw.write_block(5+2*table_pos, block3, start_pos=(5, 4), block_size=(15, 4))
        self.aw.write_addr(5+2*table_pos, meter_addr)

        # 反向有功基本误差(ABC相)
        block4 = []
        component_list = ['2', '3', '4']
        factor_list = ['1.0', '0.5L']
        I_list1 = ['Imax', '1.0Ib', '0.5Ib', '0.1Ib']
        I_list2 = ['Imax', '1.0Ib', '0.5Ib', '0.2Ib']
        for c in component_list:
            for f in factor_list:
                I_list = I_list1 if f == '1.0' else I_list2
                for I in I_list2:
                    block4 += self.ar.get_error(power_type='1',
                        component=c, factor=f, I=I)
        self.aw.write_block(5+2*table_pos, block4, start_pos=(23, 5), block_size=(24, 4))
        self.aw.write_addr(5+2*table_pos, meter_addr)

        # 无功基本误差(总)
        block5 = []
        factor_list = ('1.0', '0.5L', '0.5C', '0.25L', '0.25C')
        I_list1 = ('Imax', '1.0Ib', '0.5Ib', '0.2Ib', '0.1Ib', '0.05Ib')
        I_list2 = ('Imax', '1.0Ib', '0.5Ib', '0.2Ib', '0.1Ib')
        I_list3 = ('1.0Ib', '0.5Ib', '0.2Ib')
        for f in factor_list:
            if f == '1.0':
                I_list = I_list1
            elif f.startswith('0.5'):
                I_list = I_list2
            else:
                I_list = I_list3
            for I in I_list:
                block5 += self.ar.get_error(power_type='2',
                    component='1', factor=f, I=I)
        self.aw.write_block(14+2*table_pos, block5, start_pos=(4, 4), block_size=(22, 4))
        self.aw.write_addr(14+2*table_pos, meter_addr)

        # 无功基本误差(ABC相)
        block6 = []
        component_list = ['2', '3', '4']
        factor_list = ['1.0', '0.5L', '0.5C']
        I_list1 = ['Imax', '1.0Ib', '0.5Ib', '0.1Ib']
        I_list2 = ['Imax', '1.0Ib', '0.5Ib', '0.2Ib']
        for c in component_list:
            for f in factor_list:
                I_list = I_list1 if f == '1.0' else I_list2
                for I in I_list2:
                    block6 += self.ar.get_error(power_type='2',
                        component=c, factor=f, I=I)
        self.aw.write_block(15+2*table_pos, block6, start_pos=(4, 5), block_size=(36, 4))
        self.aw.write_addr(15+2*table_pos, meter_addr)

    def save_report(self):
        self.aw.save()


def check_meter_type(meter_addr):
    pass

if __name__ == '__main__':
    mdb_filepath = r'data/ClouMeterData_original.mdb'
    # mdb_filepath = r'd:/client/database/ClouMeterData.mdb'
    # file_path = r'./templates/'
    file_path = r'./data/1抽检原始记录A1级三相外置.docx'
    
    # meter_list = (
    #     '970000041014',
    #     '910003688788',
    #     '910003688788',
    #     '910003688788',
    #     '910003688788'
    #     )
    meter_list = []
    for i in range(5):
        meter_list.append(input(f'输入电表ID{i+1}:'))
    print('等一会儿嗷......')
    awesom_report = AwesomeReport(mdb_filepath, file_path, meter_list)
    awesom_report.report_all()
    awesom_report.save_report()


import os.path
import configparser

config = configparser.ConfigParser()
config.read('config.ini')
template_directory = config.get('install', 'template_directory')


class Document(object):
    """base class encapsulating the logic of building up a document."""
    def __init__(self, meter_addr_list):
        """building logic"""
        self.set_template()
        print(f'writing {os.path.basename(self._template)} ...')
        self.build()
        self.save()

    def build(self):
        id_data_list = [od.IdData(addr) for addr in meter_addr_list]
        self.build_deviation(id_data_list[:5])
        self.build_jidu(id_data_list[:6])
        self.build_xuliang(id_data_list[1])
        self.build_biancha(id_data_list[2])
        self.build_yizhixing(id_data_list[1:4])
        self.build_fuzaidianliu(id_data_list[4])

    def set_template(self):
        raise NotImplementedError

    def build_deviation(self, id_data_list):
        raise NotImplementedError

    def build_jidu(self, id_data_list):
        raise NotImplementedError

    def build_xuliang(self, id_data):
        raise NotImplementedError

    def build_biancha(self, id_data):
        raise NotImplementedError

    def build_yizhixing(self, id_data_list):
        raise NotImplementedError

    def build_fuzaidianliu(self, id_data):
        raise NotImplementedError

    def save(self):
        raise NotImplementedError


class OriginalRecordDocument(Document):
    def set_template(self):
        template = '1抽检原始记录A1级三相外置.docx'
        od.set_template(os.path.join(template_directory, template))

    def build_deviation(self, id_data_list):
        for i, id_data in enumerate(id_data_list):
            table_info = [(3, (4, 3), 'active', 'balanced'),
                          (3, (22, 4), 'active', 'unbalanced'),
                          (4, (4, 3), 'reversed', 'balanced'),
                          (4, (22, 4), 'reversed', 'unbalanced'),
                          (13, (3, 3), 'reactive', 'balanced'),
                          (14, (3, 4), 'reactive', 'unbalanced')]
            # print(self._id_data._meter_address, self._id_data.meter_id)
            for tid, spos, ptype, comp in table_info:
                id_data.fill(tid + 2 * i, spos)
                od.DeviationData(id_data, ptype, comp).fill(tid + 2 * i, spos)

    def build_jidu(self, id_data_list):
        for i, id_data in enumerate(id_data_list):
            id_data.fill(26, (1 + 3 * i, 0))
            od.JiduData(id_data).fill(26, (1 + 3 * i, 2))

    def build_xuliang(self, id_data):
        id_data.fill(28, (0, 1))
        od.XuliangData(id_data).fill(28, (3, 3))

    def build_biancha(self, id_data):
        id_data.fill(29, (0, 1))
        od.BianchaData(id_data).fill(29, (4, 2))

    def build_yizhixing(self, id_data_list):
        yzx_data_list = [od.YizhixingData(id_data) for id_data in id_data_list]
        for i, ele in enumerate(zip(id_data_list, yzx_data_list)):
            id_data, yzx_data = ele
            id_data.fill(30, (2 + 3 * i, 0))
            yzx_data.fill(30, (2 + 3 * i, 3))
        od.YizhixingMeanData(yzx_data_list).fill(30, (2, 6))

    def build_fuzaidianliu(self, id_data):
        id_data.fill(31, (2, 0))
        fzdl_data = od.FuzaidianliuData(id_data)
        fzdl_data.fill(31, (2, 3))
        od.FuzaidianliuAggrData(fzdl_data).fill(31, (2, 7))


if __name__ == "__main__":

    def load_meter_addr():
        import configparser as cp
        import json
        config = cp.ConfigParser()
        config.read('config.ini')
        temp = config.get('input', 'meter_addr_list')
        meter_addr_list = list(map(str, json.loads(temp)))
        return meter_addr_list

    meter_addr_list = load_meter_addr()
    Original_Record_10S(meter_addr_list)

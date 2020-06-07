import os.path
import configparser

config = configparser.ConfigParser()
config.read('config.ini')
template_directory = config.get('install', 'template_directory')


class Document(object):
    """base class encapsulating the logic of building up a document."""
    def build(self):
        self.set_template()
        self.build_deviation()
        self.build_jidu()
        self.build_xuliang()
        self.build_biancha()
        self.build_yizhixing()
        self.build_fuzaidianliu()
        self.save()

    def set_template(self):
        raise NotImplementedError

    def build_deviation(self):
        raise NotImplementedError

    def build_jidu(self):
        raise NotImplementedError

    def build_xuliang(self):
        raise NotImplementedError

    def build_biancha(self):
        raise NotImplementedError

    def build_yizhixing(self):
        raise NotImplementedError

    def build_fuzaidianliu(self):
        raise NotImplementedError

    def save(self):
        raise NotImplementedError


class OriginalRecordDocument(Document):
    def __init__(self, meter_addr_list):
        self._od = __import__('original_data')
        self._id_data_list = [
            self._od.IdData(addr) for addr in meter_addr_list
        ]

    def set_template(self, template):
        # template = '1抽检原始记录A1级三相外置.docx'
        print(f'writing {os.path.basename(template)} ...')
        self._od.set_template(os.path.join(template_directory, template))

    def build_deviation(self):
        for i, id_data in enumerate(self._id_data_list[:5]):
            table_info = [(3, (4, 3), 'active', 'balanced'),
                          (3, (None, 4), 'active', 'unbalanced'),
                          (4, (4, 3), 'reversed', 'balanced'),
                          (4, (None, 4), 'reversed', 'unbalanced'),
                          (13, (3, 3), 'reactive', 'balanced'),
                          (14, (3, 4), 'reactive', 'unbalanced')]
            # print(self._id_data._meter_address, self._id_data.meter_id)
            for tid in (3, 4, 13, 14):
                id_data.fill(tid + 2 * i, (0, 1))
            prev = None
            for tid, spos, ptype, comp in table_info:
                if spos is None:
                    spos = 7 + prev.content.shape
                prev = self._od.DeviationData(id_data, ptype, comp)
                prev.fill(tid + 2 * i, spos)

    def build_jidu(self):
        for i, id_data in enumerate(self._id_data_list[:6]):
            id_data.fill(26, (1 + 3 * i, 0))
            self._od.JiduData(id_data).fill(26, (1 + 3 * i, 2))

    def build_xuliang(self):
        id_data = self._id_data_list[1]
        id_data.fill(28, (0, 1))
        self._od.XuliangData(id_data).fill(28, (3, 3))

    def build_biancha(self):
        id_data = self._id_data_list[2]
        id_data.fill(29, (0, 1))
        self._od.BianchaData(id_data).fill(29, (4, 2))

    def build_yizhixing(self):
        id_data_list = self._id_data_list[1:4]
        yzx_data_list = [
            self._od.YizhixingData(id_data) for id_data in id_data_list
        ]
        for i, ele in enumerate(zip(id_data_list, yzx_data_list)):
            id_data, yzx_data = ele
            id_data.fill(30, (2 + 3 * i, 0))
            yzx_data.fill(30, (2 + 3 * i, 3))
        self._od.YizhixingMeanData(yzx_data_list).fill(30, (2, 6))

    def build_fuzaidianliu(self):
        id_data = self._id_data_list[4]
        id_data.fill(31, (2, 0))
        fzdl_data = self._od.FuzaidianliuData(id_data)
        fzdl_data.fill(31, (2, 3))
        self._od.FuzaidianliuAggrData(fzdl_data).fill(31, (2, 7))

    def save(self):
        self._od.save()


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
    # test = OriginalRecordDocument(meter_addr_list)
    # test.set_template('1抽检原始记录A1级三相外置.docx')
    # test.build_deviation()
    # test.save()
    OriginalRecordDocument(meter_addr_list).build()

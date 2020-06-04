from writer import Writer
from data import DeviationData, JiduData, XuliangData, BianchaData, \
        YizhixingData, YizhixingMeanData, FuzaidianliuData, \
        FuzaidianliuAggrData


class Block(object):

    _writer = Writer(r'data/1抽检原始记录A1级三相外置.docx')

    def __init__(self, position, data):
        # position: (table_id, (start_posx, start_posy))
        self._table_id, self._start_pos = position
        self._data = data

    def fill(self):
        try:
            self._writer.write(self._table_id, self._data.read(),
                               self._start_pos)
        except Exception as e:
            self._writer.save()
            print("write aborted.")
            raise e

    @classmethod
    def save(cls):
        cls._writer.save()


class FormAbstractClass(object):
    def __init__(self, id_data):
        self._id_data = id_data

    def fill(self):
        raise NotImplementedError

    @classmethod
    def save(cls):
        Block.save()


class DeviationForm(FormAbstractClass):

    _count = 0

    def fill(self):
        table_info = [(3, (4, 3), 'active', 'balanced'),
                      (3, (22, 4), 'active', 'unbalanced'),
                      (4, (4, 3), 'reversed', 'balanced'),
                      (4, (22, 4), 'reversed', 'unbalanced'),
                      (13, (3, 3), 'reactive', 'balanced'),
                      (14, (3, 4), 'reactive', 'unbalanced')]
        # print(self._id_data._meter_address, self._id_data.meter_id)
        for tid, spos, ptype, comp in table_info:
            # print(tid, spos, ptype, comp)
            data = self._id_data
            Block((tid + 2 * DeviationForm._count, (0, 1)), data).fill()
            data = DeviationData(self._id_data, ptype, comp)
            Block((tid + 2 * DeviationForm._count, spos), data).fill()
        DeviationForm._count += 1


class JiduForm(FormAbstractClass):

    _count = 0

    def fill(self):
        data = JiduData(self._id_data)
        Block((26, (1 + 3 * self._count, 0)), self._id_data).fill()
        Block((26, (1 + 3 * self._count, 2)), data).fill()
        JiduForm._count += 1


class XuliangForm(FormAbstractClass):
    def fill(self):
        data = XuliangData(self._id_data)
        Block((28, (3, 3)), data).fill()
        Block((28, (0, 1)), self._id_data).fill()


class BianchaForm(FormAbstractClass):
    def fill(self):
        data = BianchaData(self._id_data)
        Block((29, (4, 2)), data).fill()
        Block((29, (0, 1)), self._id_data).fill()


class YizhixingForm(FormAbstractClass):

    _count = 0
    _yzx_list = []

    def fill(self):
        data = YizhixingData(self._id_data)
        Block((30, (2 + 3 * self._count, 3)), data).fill()
        Block((30, (2 + 3 * self._count, 0)), self._id_data).fill()
        self._yzx_list.append(data)
        YizhixingForm._count += 1
        if self._count == 3:
            data = YizhixingMeanData(self._yzx_list)
            Block((30, (2, 6)), data).fill()


class FuzaidianliuForm(FormAbstractClass):
    def fill(self):
        data = FuzaidianliuData(self._id_data)
        Block((31, (2, 3)), data).fill()
        data = FuzaidianliuAggrData(data)
        Block((31, (2, 7)), data).fill()
        Block((31, (2, 0)), self._id_data).fill()


if __name__ == "__main__":
    from data import IdData
    import configparser as cp
    import json
    config = cp.ConfigParser()
    config.read('config.ini')
    temp = config.get('input', 'meter_addr_list')
    meter_addr_list = list(map(str, json.loads(temp)))

    id_data_list = [IdData(addr) for addr in meter_addr_list]
    print(meter_addr_list)
    print('wait please.')
    Block((29, (0, 1)), id_data_list[0]).fill()
    Block.save()
    print('done.')

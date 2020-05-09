from writer import Writer
from data import DeviationData


class Block(object):

    _writer = Writer(r'data/1抽检原始记录A1级三相外置.docx')

    def __init__(self, position, data):
        # position: (table_id, (start_posx, start_posy))
        self._table_id, self._start_pos = position
        self._data = data

    def fill(self):
        _writer.write(self._table_id, self._data.read(), self._start_pos)


class DeviationForm:

    _count = 0

    def __init__(self, id_data):
        self._id_data = id_data

    def fill(self):
        # count 没实现
        blist = []
        blist.append(Block((3+count, (0,1)), self._id_data))
        blist.append(Block((3+count, (4,3)), DeviationData(self._id_data, 'active', 'balanced')))
        blist.append(Block((3+count, (18,4)), DeviationData(self._id_data, 'active', 'unbalanced')))
        for block in blist:
            block.fill()
        self._count += 1
        
if __name__ == "__main__":
    meter_addr_list = ['xxx','xxx','xxx','xxx','xxx','xxx']
    iddata_list = [IdData(meter_addr) for meter_addr in meter_addr_list]
    for iddata in iddata_list:
        DeviationForm(iddata).fill()


import data as dt
from writer_singleton import writer


def save_file():
    writer.save()


def set_template(file_path):
    writer.set_template(file_path)


class Block(dt.DataAbstractClass):
    def fill(self, tab_pos, start_cell):
        data = super().read()
        writer.write(tab_pos, data, start_cell)


class IdBlock(Block, dt.IdData):
    pass


class DeviationBlock(Block, dt.DeviationData):
    pass


class JiduBlock(Block, dt.JiduData):
    pass


class XuliangBlock(Block, dt.XuliangData):
    pass


class BianchaBlock(Block, dt.BianchaData):
    pass


class YizhixingBlock(Block, dt.YizhixingData):
    pass


class YizhixingMeanBlock(Block, dt.YizhixingMeanData):
    pass


class FuzaidianliuBlock(Block, dt.FuzaidianliuData):
    pass


class FuzaidianliuAggrData(Block, dt.FuzaidianliuAggrData):
    pass


if __name__ == "__main__":
    import configparser as cp
    import json
    config = cp.ConfigParser()
    config.read('config.ini')
    temp = config.get('input', 'meter_addr_list')
    meter_addr_list = list(map(str, json.loads(temp)))
    id_data_list = [dt.IdData(addr) for addr in meter_addr_list]
    JiduBlock(id_data_list[0]).fill(26, (1, 2))

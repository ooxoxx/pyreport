import data as dt
import writer


class Block(dt.DataAbstractClass):
    """Docstring for Block. """

    _writer = None

    def __init__(self, *args, **kargs):
        """TODO: to be defined.

        :*args: TODO
        :**kargs: TODO

        """
        super().__init__(*args, **kargs)
        if self._writer is None:
            template_file_path = "./data/1抽检原始记录A1级三相外置.docx"
            Block._writer = writer.Writer(template_file_path)

    def fill(self, tab_pos, start_cell):
        """TODO: Docstring for fill.

        :tab_pos: TODO
        :start_cell: TODO
        :returns: TODO

        """
        data = super().read()
        self._writer.write(tab_pos, data, start_cell)
        self._writer.save()


class IdBlock(Block, dt.IdData):
    pass


class JiduBlock(Block, dt.JiduData):
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

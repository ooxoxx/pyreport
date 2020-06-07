import original_data as od
import numpy as np


class Result(od.Data):
    def __init__(self, data):
        self._data = data

    def read(self):
        return self._data.content[:, -1].reshape(-1, 1)


class DeviationResult(Result):
    pass


class JiduResult(Result):
    def read(self):
        res = self._data.content[-1, :].reshape(1, -1)
        res = np.c_[res[:, :-1], np.array([['±0.03']]), res[:, -1]]
        return res


class XuliangResult(Result):
    pass


class BianchaResult(Result):
    pass


class YizhixingResult(Result):
    def read(self):
        return self._data.content[:, -2].reshape(-1, 1)


class FuzaidianliuResult(Result):
    pass


if __name__ == "__main__":
    import configparser as cp
    import json
    config = cp.ConfigParser()
    config.read('config.ini')
    temp = config.get('input', 'meter_addr_list')
    meter_addr_list = list(map(str, json.loads(temp)))

    id_data_list = [od.IdData(addr) for addr in meter_addr_list]
    # print(DeviationResult(id_data_list[0], 'active', 'balanced').read())
    print(JiduResult(od.JiduData(id_data_list[0])).read())

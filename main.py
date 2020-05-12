from form import DeviationForm, JiduForm, XuliangForm, \
        BianchaForm, YizhixingForm, FuzaidianliuForm
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

for i, id_data in enumerate(id_data_list):
    print(f'{id_data.read()[0][0]} processing.')
    if i < 5:
        DeviationForm(id_data).fill()
    JiduForm(id_data).fill()
XuliangForm(id_data_list[1]).fill()
BianchaForm(id_data_list[2]).fill()

YizhixingForm(id_data_list[1]).fill()
YizhixingForm(id_data_list[2]).fill()
YizhixingForm(id_data_list[3]).fill()

FuzaidianliuForm(id_data_list[4]).fill()

FuzaidianliuForm.save()
print('done.')

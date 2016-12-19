# coding=utf-8

"""

读取原始的eos数据，生成字典文件：
{ "bom编码"：["EOS DCP计划"，"EOS DCP实际"，"EOS公告上网实际", "EOS公告上网计划", "EOL DCP实际", "EOL 计划"]}
将生成的字典通过pickle 序列化保存，供查询时使用
"""

import pickle


def pickle_data(eos_data):
    with open(r".\eos_data\eos-data", 'wb') as f:
        pickle.dump(eos_data, f)


def get_data(path):
    f = open(path, encoding='utf-8')
    return f.readlines()


eos_data_dict = {}
PATH = r".\eos_data\eox-1117.txt"

lines = get_data(PATH)

for line in lines:
    # print(line)
    field = line.replace('\n', '').split(',')
    print (field[2])
    print(field[5:7] + field[10:12] + field[7:9] )
    eos_data_dict[field[2]] = field[5:7] + field[10:12] + field[7:9]

print(eos_data_dict["0231A84Q"])
print(eos_data_dict)
print(len(eos_data_dict))

pickle_data(eos_data_dict)

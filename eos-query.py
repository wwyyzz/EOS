# -*- coding: UTF-8 -*-
"""
write by navy
2016-12-12
"""

import pickle
import re
import sqlite3
import xlrd
import xlwt
import os


def get_eos_data():
    """
    功能：通过pickle获取eos_data字典数据
    返回：eos_data字典数据

    """
    eos_data = {}
    with open(r".\eos_data\eos-data", 'rb') as f:
        eos_data = pickle.load(f)

    print("共有记录：" + str(len(eos_data)))
    return eos_data


def get_device_type(version_info):
    """
    功能：查找display version获取设备型号
    通过正则，匹配设备型号，非H3C型号可不考虑
    <AHTLWA01-C1> display version
     H3C Comware Platform Software
     Comware Software, Version 5.20.106, Release 3303P20
     Copyright (c) 2004-2015 Hangzhou H3C Tech. Co., Ltd. All rights reserved.
     H3C SR6608 uptime is 9 weeks, 4 days, 2 hours, 17 minutes
     Slot 0: RPE-X1 uptime is 9 weeks, 4 days, 2 hours, 17 minutes
    返回：单台设备的型号、板卡名称、序列号
    """
    pattern_h3c = re.compile(r'\n(H3C\s.*)\suptime\sis')
    pattern_other = re.compile(r'\n(.*)\suptime\sis')

    h3c_match = re.search(pattern_h3c, version_info)
    if h3c_match is not None:
        device_type = h3c_match.group(1)
    else:
        other_match = re.search(pattern_other, version_info)
        if other_match is not None:
            device_type = other_match.group(1)
        else:
            device_type = "unknown device"

    return device_type


def get_device_info(version_info, manu_info):
    """
    功能：获取设备信息信息
    通过正则，查找所有的DEVICE_NAME、DEVICE_SERIAL_NUMBER信息放入列表中
    Slot 3:
    DEVICE_NAME:FIP-200
    DEVICE_SERIAL_NUMBER:210231A763B103000062
    MAC_ADDRESS:0023-89A6-B3F4
    MANUFACTURING_DATE:2010-03-15
    VENDOR_NAME:H3C
    返回：单台设备的型号、板卡名称、序列号
    :param version_info:
    :param manu_info:
    :return: 单台设备的型号、板卡名称、序列号
    """

    device_type = get_device_type(version_info)

    pattern_device_name = re.compile(r'DEVICE[_\s]NAME\s*:\s*(.+)\n')
    pattern_device_sn = re.compile(r'DEVICE[_\s]SERIAL[_\s]NUMBER\s*:\s*(\S+)\n')

    if device_type[0:3] == 'H3C':
        device_name = re.findall(pattern_device_name, manu_info)
        device_sn = re.findall(pattern_device_sn, manu_info)
        manu_info_list = [[a, b] for a, b in zip(device_name, device_sn)]
    else:
        manu_info_list = [['unknown', '   unknown']]

    device_info = [device_type, manu_info_list]
    # print("device_info ============")
    # print(device_info)
    return device_info


def get_summary_of_device(summary_list):
    """
    功能：获取板块的汇总统计信息
    将输入列表保存到数据库中，使用SQL进行数据汇总统计
    返回：基于型号、板卡的数量汇总
    """

    #    conn = sqlite3.connect(r'D:\1-MY\2-Code\Python\EOS\device.db')
    conn = sqlite3.connect(r'C:\MyCode\EOS\device.db')
    c = conn.cursor()

    c.execute("delete from DEVICE;")
    c.execute("update sqlite_sequence SET seq = 0 where name ='DEVICE';")

    print("write_db ...................")

    for device in summary_list:
        for num in range(len(device[1])):
            c.execute("INSERT INTO DEVICE (device_type, module_type, module_sn, bom) VALUES (?, ?, ?, ?  )",
                      [device[0], device[1][num][0], device[1][num][1], device[1][num][1][2:10]])

    conn.commit()
    c.execute(
        "SELECT device_type, module_type, bom, COUNT(module_type) from DEVICE GROUP BY device_type, module_type ORDER BY device_type, module_type")
    summary_of_device = c.fetchall()
    c.close()

    result_list = []
    for moudle in summary_of_device:
        moudle_list = list(moudle)
        #    result_list.append((list)eos_data.get(moudle[2]))
        eos_query = eos_data.get(moudle[2])
        if eos_query is not None:
            result = moudle_list + eos_query
        else:
            result = moudle_list + ['N/A'] * 7

        result_list.append(result)

    return result_list


def get_devices_all(filename):
    """
    功能：获取无重复的设备型号、板卡型号和序列号sn
    读取指定文件,调用get_device_info 对 display version 和 display device manuinfo 信息进行解析，获取设备型号、版本型号和序列号sn
    对于重复的信息进行去重
    返回：无重复数据的列表
    ['H3C MSR30-11', [['MSR 30-11', '210235A274B081000034'], ['RT-XMIM-24FSW', '210231A77BB081000272']]]
    """
    devices_all = []

    book = xlrd.open_workbook(filename)
    sh = book.sheet_by_index(0)
    for i in range(1, sh.nrows - 1, 2):
        device = get_device_info(sh.cell_value(i, 3), sh.cell_value(i + 1, 3))
        devices_all.append(device)

    devices_all.sort()

    devices_all_duplicate_removal = []
    for device in devices_all:
        if device not in devices_all_duplicate_removal:
            devices_all_duplicate_removal.append(device)

    return devices_all_duplicate_removal


def write_xls(result, filename):
    """
    输入: 结果列表,保存文件名
    将汇总和查询结果写入到输出xls文件中
    """

    book = xlwt.Workbook()
    sheet1 = book.add_sheet(u'sheet1', cell_overwrite_ok=True)

    row0 = [u'型号', u'板卡类型', u'BOM编码', u'数量', u"产品线", u"所属PDT", u"EOM DCP实际时间", u"EOS DCP计划时间", u"EOS DCP实际时间",
            u"EOS公告上网实际时间", u"EOS公告上网计划时间"]
    for i in range(0, len(row0)):
        sheet1.col(i).width = 256 * 20
        sheet1.write(0, i, row0[i])

    n = 1
    for line in result:
        for i in range(0, 11):
            sheet1.write(n, i, line[i])
        n += 1

    book.save(filename)


'''
主程序
从指定的目录中读取需要查询的xls文件，逐个文件进行解析，将汇总结果保存到输出文件夹中
'''

INPUT_DIR = r'.\\H3C-display\\'
OUTPUT_DIR = r'.\\output\\'
eos_data = get_eos_data()

# files = os.listdir(r'.\\H3C-display\\')
files = os.listdir(r'.\\test-input\\')

for filename in files:
    PATH = INPUT_DIR + filename
    output_filename = OUTPUT_DIR + filename.split('.')[0] + r'-summary.xls'
    print(r"保存文件名为:" + output_filename)
    devices_all_duplicate_removal = get_devices_all(PATH)
    summary_of_device = get_summary_of_device(devices_all_duplicate_removal)
    write_xls(summary_of_device, output_filename)

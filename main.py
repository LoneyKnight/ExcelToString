import sys
import os
import re
import ToLua
import xlrd
import json
import time
import csv


def LoadFiles(files):
    root = {}
    for file in files:
        print("------开始读取%s------" % file)
        # if file == 'string.csv':
        #     root = load_csv(file, root)
        # else:
        #     continue
        info = xlrd.open_workbook(file)
        # table = data.sheet_by_name('main@')
        # 读取每个sheet
        data_sheet = info.sheets()
        for table in data_sheet:
            if table.name[-1] == '#':
                continue
            idCol, cnCol, nameCol, descCol = 0, 0, 0, 0
            for i in range(table.ncols):
                if '字符串' in table.cell_value(0, i):
                    idCol = i - 1
                    cnCol = i
                    enCol = i + 1
                for j in range(3, table.nrows):
                    if cnCol > 0:
                        if table.cell_value(j, cnCol) == '':
                            continue
                        text = ParseString(str(table.cell_value(j, cnCol)))
                        root[table.cell_value(j, idCol)] = {'cn': text}
    return root


def load_csv(file, root):
    with open(file, encoding='utf-8-sig', mode='r') as f:
        # 读取string.csv
        csv_info = csv.DictReader(f)
        for row in csv_info:
            root[str(row['id'])] = {'cn': row['cn']}
        return root


def ToJson(info, output_dir):
    with open(r'%s/string.json' % output_dir, 'w', encoding='UTF-8') as fileobject:
        fileobject.write(json.dumps(info, indent=4, ensure_ascii=False))
    print("------写入string.json完毕------")


# 将字符串转换为lua格式
def saveLua(info, output_dir):
    json_text = json.dumps(info, indent=4, ensure_ascii=False)
    lua_text = ToLua.str_to_lua_table(json_text)
    # 保存lua文件
    with open(r'%s/stringConfig.lua' % output_dir, 'w', encoding='UTF-8') as fileobject:
        fileobject.write(lua_text)
    print("------写入string.lua完毕------")


def ToTxt(info, output_dir):
    with open(r'%s/string.txt' % output_dir, 'w', encoding='UTF-8') as fileobject:
        for key, value in info.items():
            fileobject.write("%s,%s\n" % (key, value['cn']))


# 解析字符串中包含{}的字符串
def ParseString(string):
    # 正则匹配{}和之间的内容
    pattern = re.compile(r'{(.*?)}')
    result = pattern.findall(string)
    # 如果没有匹配到{}，则直接返回
    if len(result) == 0:
        return string
    # 如果匹配到{}，则进行替换
    for i in range(len(result)):
        if "name" in result[i]:
            string = string.replace("{%s}" % result[i], "<color=#a75c16>{%s}</color>" % result[i])
    # 如果匹配到<ts>，则进行替换
    if "<ts>" in string:
        string = string.replace("<ts>", "<color=#a75c16>")
        string = string.replace("</ts>", "</color>")
    return string


if __name__ == '__main__':
    config = json.load(open('./files.json', 'r'))
    files = config['files']
    outputDir = config['outputDir']
    try:
        data = LoadFiles(files)
        # ToJson(data, outputDir)
        data = load_csv('string.csv', data)
        saveLua(data, outputDir)
        if 'isOutputTxt' in config and config['isOutputTxt'] == 1:
            ToTxt(data, os.getcwd())
            print('------输出txt完毕,路径为:%s------' % os.getcwd())
    except Exception as e:
        print(e)
        print("↑出现问题,请查看配置或者脚本是否正确,按任意键关闭")
        input()
    print("**** 3秒后关闭窗口 ****")
    time.sleep(3)

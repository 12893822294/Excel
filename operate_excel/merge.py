###############################################
# author: wuxingkai
# date: 2021-04-03
# 按模板excel的表头合并文件，如果没有指定模板文件，则第一个excel就为模板
###############################################

import os
from operate_excel import ReadExcel, WriteExcel


def get_files(path, file_suffix='.xlsx', result=None):
    """
    深度遍历获取一个目录下面所有符合file_suffix的文件
    :param path: 获取改目录下的文件
    :param file_suffix: 文件名的关键字，一般是后缀
    :param result:
    :return:
    """
    if not path:
        raise Exception("get_files Exception :{}".format('请指定一个目录'))
    if not result:
        result = []
    if not os.path.isdir(path) and file_suffix in path:
        return [path]
    files = [path + '\\' + _ for _ in os.listdir(path)]
    for file in files:
        if os.path.isdir(file):
            # 如果是文件夹，继续遍历里面的文件
            result += get_files(file, file_suffix)
        else:
            if file_suffix in file:
                # 如果是所要的文件，就添加到file_path_list
                result.append(file)
    return result


def run(file_path, demo_file=None, file_suffix='xlsx'):
    """
    开始合并
    :param file_path:
    :param demo_file:
    :param file_suffix:
    :return:
    """
    write_result = {}
    other_result = {}
    files = get_files(file_path, file_suffix=file_suffix)
    other_head = input('按哪一列分开合并>>>')
    if other_head:
        other_value = input('这一列的什么值单独拿出来>>>')
    else:
        other_value = ''
    # merge_word = {other_head: other_value}
    if not demo_file:
        demo_file = files[0]
    excel = ReadExcel(file_path=demo_file)
    heads = excel.read_table_head()
    for file in files:
        excel = ReadExcel(file_path=file)
        # a = excel.read_sheets_data()
        read_result = excel.read_sheets_data(heads=heads)
        for sheet, sheet_result in read_result.items():
            if sheet not in write_result:
                write_result[sheet] = []
            if sheet not in other_result:
                other_result[sheet] = []
            for row_dict in sheet_result:
                if other_head and other_head not in row_dict:
                    raise Exception("没有这一列:{}".format(other_head))
                if other_head and str(row_dict[other_head]) == other_value:
                    other_result[sheet].append(row_dict)
                else:
                    write_result[sheet].append(row_dict)
    if write_result:
        save_name = input('保存\"{}\"不为\"{}\"的文件名(默认：\"{}.xlsx\")>>>').format(other_head, other_value, "Undefined")
        save_name = save_name if save_name else "Undefined.xlsx"
        WriteExcel(write_result=write_result).write_sheets(
            save_name=save_name
        )
    if other_result:
        save_name = input('保存\"{}\"等于\"{}\"的文件名(默认：\"{}.xlsx\")>>>'.format(other_head, other_value, other_value))
        save_name = save_name if save_name else other_value
        WriteExcel(write_result=other_result).write_sheets(
            save_name=save_name
        )


if __name__ == '__main__':
    # run('')
    print(get_files(r'D:\py_project', file_suffix='.py'))

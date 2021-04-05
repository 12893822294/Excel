接口：merge.py
入口参数：
input输入一个目录"file_path"，合并该目录下的文件
input输入一个模板文件，会按照该文件的表头去合并
input输入文件的关键字（默认为xlsx），会合并"file_path"目录下所有包含关键字的文件
input输入要按哪一列合并的表头"other_head"，以及这一列的值"other_value", 一般用于In_Doubt分'Y'和'N'

把所有文件按不同的sheet保存在一个excel，如果指定的 
会把这一列指定的值合并一份，不为输入值的合并一份，

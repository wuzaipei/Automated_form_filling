import xlwt, xlrd
from datetime import date, datetime

# 打开excel文件，创建一个workbook对象,book对象也就是fruits.xlsx文件,表含有sheet名
workbook = xlrd.open_workbook(r'./surfaceFile/表2.xlsx')
print(workbook.sheet_names())  # 得到一个列表里面的元素就是sheet的名字
# 上面执行的结果为['Sheet1', 'Sheet2']
print(workbook.sheets())  # 得到的是一个列表里面的元素就是每一个sheet对象
# 上面执行的结果为[<xlrd.sheet.Sheet object at 0x0000019BB9D915C0>, <xlrd.sheet.Sheet object at 0x0000019BB9D91128>]

sheet_name = workbook.sheet_names()[0]  # 从零开始，取第一个sheet的名字
sheet_obj = workbook.sheets()[0]  # 从零开始取第一个sheet对象
print(sheet_name, sheet_obj)
# 上面执行的结果为：Sheet1 <xlrd.sheet.Sheet object at 0x000001E620A7CBE0>

# 根据sheet索引或者名称获取sheet内容
rsheet = workbook.sheet_by_index(0)  # 取第一个工作簿根据索引
rsheet_name = workbook.sheet_by_name(sheet_name)  # 根据sheet的名字取第一个工作簿
print('rsheet_index', rsheet)
print('rsheet_name', rsheet_name)

# 获取总行数,列数和名字根据sheet的内容也就是上面的rsheet或者rsheet_name
print(rsheet.nrows, rsheet.ncols, rsheet.name)
rows = rsheet.nrows
# 获取总列数
cols = rsheet.ncols
# sheet名称
sheet_name = rsheet.name

# 获取整行和整列的值
rows2_values = rsheet.row_values(1)  # 获取第二行内容，得到的是一个列表
cols3_values = rsheet.col_values(2)  # 获取第三列内容，得到的是一个列表
print(rows2_values, cols3_values)
# ['小杰', 23.0, 33919.0, '键盘', '朋友'] ['出生日期', 33919.0, 33920.0, 33921.0, 33922.0, 33923.0, '暂无']

# 通过cell的位置坐标取得cell值的几种方式
print('获取第二行第一列的值', rsheet.cell(1, 0).value)
print('获取第二行第一列的值', rsheet.cell_value(1, 0))
print('获取第二行第一列的值', rsheet.row(1)[0].value)

# 获取单元格内容的数据类型
print(rsheet.cell(1, 0).ctype)
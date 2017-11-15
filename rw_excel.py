# 所需包xlrd(read xls),和xlwt(write xls)
import xlrd
import xlwt as ExcelWrite
# sheet编号，行号，列号都是从索引0开始

# 打开Excel
data = xlrd.open_workbook('test.xlsx')
# 读取第一个sheet
sheet = data.sheets()[0]
# 获取第i行数据
rows = sheet.row_values(0)
# 获取第i列数据
my_queues = sheet.col_values(0)
for rowsData in rows:
    # 此处若不加上int数字会默认加上.0的格式
    # 在Python2.x中还可以使用long()
    print(int(rowsData))
print (my_queues)

def mywriteXLS(cols_one, clols_two, file_name):
    xls = ExcelWrite.Workbook()
    sheet = xls.add_sheet("Sheet1", cell_overwrite_ok=True)
    i = 0
    for row in clos_one:
        for row_one in clos_two:
            sheet.write(i, 0, row)
            sheet.write(i, 1, row_one)
            i = i+1
    xls.save(file_name)


def readXLS(file_name):
    one_lists = []
    two_lists = []
    one = ''
    two = ''
    data = xlrd.open_workbook(file_name)
    sheet = data.sheets()[9]
    one_queues = sheet.col_values(0)
    two_queues = sheet.col_values(1)
    for one in one_queues:
        if one != '':
            one_lists.append(one)
    for two in two_queues:
        if two != '':
            two_lists.append(two)
    mywriteXLS(one_lists, two_lists, 'file9.xls')


if __name__ == "__main__":
    readXLS("你需要阅读的EXCEL文件")

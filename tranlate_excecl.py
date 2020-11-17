'''
翻译excel
1 读入 excel
2 翻译
3 写入excel
'''
import xlrd
from googletrans import Translator
from xlutils.copy import copy as xl_copy
import xlwt

class Tran():
    def __init__(self):
        pass
    def read_exel(self):
        pass
    def translate(self):
        pass
    def write_excel(self):
        pass

'''
1 读入excel
'''
def read_excecl(name):
    # 打开刚才我们写入的 test_w.xls 文件
    wb = xlrd.open_workbook(name)

    # 获取并打印 sheet 数量
    # print( "sheet 数量:", wb.nsheets)

    # 获取并打印 sheet 名称
    # print( "sheet 名称:", wb.sheet_names())

    # 根据 sheet 索引获取内容 第一个sheet
    # sh1 = wb.sheet_by_index(0)
    # 或者也可根据 sheet 名称获取内容
    # sh = wb.sheet_by_name('成绩')
    sh1 = wb.sheet_by_name('销售费用')
    # 获取并打印该 sheet 行数和列数
    # print( u"sheet %s 共 %d 行 %d 列" % (sh1.name, sh1.nrows, sh1.ncols))

    # 获取并打印某个单元格的值
    # print( "第一行第1列的值为:", sh1.cell_value(0, 0))

    # 获取整行或整列的值
    rows = sh1.row_values(0) # 获取第一行内容
    cols = sh1.col_values(4) # 获取第一列内容
    # col_2 = sh1.col_values(1) # 获取第二列内容
    res = [cols]
    print(res)
    # print(len(res))
    return res
'''
tra
'''

def translate(datas):
    res = []
    translator = Translator(service_urls=[ 'translate.google.cn'])
    # 如果可以上外网，还可添加 'translate.google.com' 等
    # trans=translator.translate("hello", src='en', dest='zh-cn').text
    # tra = translator.translate('안녕하세요.', dest='ja').text

    # 多列的情况
    len_cols = len(datas)
    for col in range(len_cols):
        data_col = datas[col]
        res_tmp = []
        for data in data_col:
            if data !='':
                tmp = translator.translate(data, src='zh-cn', dest = 'en').text
                res_tmp.append(tmp)

        res.append(res_tmp)

    print('translate has done : ',res)
    return res

def write_to_excecl(name, data):
    wb = xlrd.open_workbook(name)
    # sh1 = wb.sheet_by_index(0)
    wb_copy = xl_copy(wb)
    trans_sheet = wb_copy.add_sheet('trans_sheet')

    # 多列的情况
    len_cols = len(data)
    for col in range(len_cols):
        data_col = data[col]
        row = len(data_col)
        for i in range(row):
            trans_sheet.write(i, col, data_col[i])
        # for data_value in data_col:
        #     trans_sheet.write(i, col, data[i])

    wb_copy.save('2.xls')
    print('write has done ')


if __name__ == '__main__':
    name = '1.xls'
    data = read_excecl(name)
    res = translate(data)
    write_to_excecl(name, res)


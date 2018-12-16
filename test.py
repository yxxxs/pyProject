import xlrd
from collections import Counter

def check_header(sheet_header):
    onrack = ['任务批次号', '业务单据号', '货主', '任务批次类型', '任务批次子类型', '任务状态','作业类型','容器','优先级','创建时间','完成时间','来源库区','目的库区','操作人','商品条码','商品名称','批次编码','生产日期','失效日期','来源库位','目标库位','计划数量','实际数量','差异数量']
    if sheet_header == onrack:
        return('onrack')
    return('shiiit')

def onrack_workload():
    ol, rackhigh_fcl, high_part, low_fcl, low_part = {}
    for rx in range(1,sh.nrows):
        if sh.cell_value(rx, 3) == '收货' and sh.cell_value(rx, 4) == '上架' and sh.cell_value(rx, 5) == '已完成':
            if





book = xlrd.open_workbook('d:\\test.xls')
sh = book.sheets()[0]

sh_header = []
for cx in range(sh.ncols):
    sh_header.append(str(sh.cell_value(0,cx)))

print(check_header(sh_header))

# print( check_header(['任务批次号', '业务单据号', '货主', '任务批次类型', '任务批次子类型', '任务状态','作业类型','容器','优先级','创建时间','完成时间','来源库区','目的库区','操作人','商品条码','商品名称','批次编码','生产日期','失效日期','来源库位','目标库位','计划数量','实际数量','差异数量']))



# z=dict(Counter(x)+Counter(y))
# print("The number of worksheets is", book.nsheets)
# print("Worksheet name(s):", book.sheet_names())
# # 打开工作表(三种方法)
# sh = book.sheet_by_index(0)
# sh = book.sheets()[0]
# sh = book.sheet_by_name('sheet1')
#
# # 操作行列和单元格
# print sh.name, sh.nrows, sh.ncols
# print "Cell D30 is", sh.cell_value(rowx=29, colx=3)
# print "Cell D30 is", sh.cell(29,3).value
#
# # 循环
# for rx in range(sh.nrows):
#     print sh.row(rx)
# # Refer to docs for more details.
# # Feedback on API is welcomed.

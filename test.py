import xlrd
import pprint
import xlwt

# 字典工具函数
def dict_val(key, dict):
    if key in dict:
        return dict[key]
    else:
        return ''
def dict_count(key, dict):
    if key in dict:
        dict[key] += 1
    else:
        dict[key] = 1
    return dict
def dict_add(key, dict, value):
    if key in dict:
        dict[key] += value
    else:
        dict[key] = value
    return dict

# 表头识别函数
def check_header(sheet_header):
    onrack = ['任务批次号', '业务单据号', '货主', '任务批次类型', '任务批次子类型', '任务状态','作业类型','容器','优先级','创建时间','完成时间','来源库区','目的库区','操作人','商品条码','商品名称','批次编码','生产日期','失效日期','来源库位','目标库位','计划数量','实际数量','差异数量']
    lights = ['日期', '拣货员工号', '拣货员', '时效', '箱数', '行数', '件数', '工时']
    if sheet_header == onrack:
        return('onrack')
    if sheet_header == lights:
        return ('lights')
    return('shit,match nothing')

# 解析上架表中的任务函数
def onrack_workload():
    row_fcl_high, row_fcl_low, row_part_high, row_part_low = {}, {}, {}, {}
    pcs_fcl_high, pcs_fcl_low, pcs_part_high, pcs_part_low = {}, {}, {}, {}
    onrack_shit = []

    for rx in range(1,sh.nrows):
        name = sh.cell_value(rx, 13)
        pcs = sh.cell_value(rx, 22)

        flag_fcl_part = sh.cell_value(rx, 13)[0:6]
        flag_high_low = int(str(sh.cell_value(rx, 20))[-1])

        if sh.cell_value(rx, 3) == '收货' and sh.cell_value(rx, 4) == '上架' and sh.cell_value(rx, 5) == '已完成':
            if flag_fcl_part in ['CC-整箱-']:
                if flag_high_low > 2:
                    dict_count(name, row_fcl_high)
                    dict_add(name, pcs_fcl_high, pcs)
                elif flag_high_low <= 2:
                    dict_count(name, row_fcl_low)
                    dict_add(name, pcs_fcl_low, pcs)
                else:
                    onrack_shit.append(name + '|' + 'row' + str(rx + 1) + '|上架，不能区分高低')
            elif flag_fcl_part in ['YW-2库-', 'CC-2库-']:
                if flag_high_low > 2:
                    dict_count(name, row_part_high)
                    dict_add(name, pcs_part_high, pcs)
                elif flag_high_low <= 2:
                    dict_count(name, row_part_low)
                    dict_add(name, pcs_part_low, pcs)
                else:
                    onrack_shit.append(name + '|' + 'row' + str(rx + 1) + '|上架，不能区分高低')
            else:
                onrack_shit.append(name + '|' + 'row' + str(rx + 1) + '|上架，不能区分整散')
        else:
            onrack_shit.append(name + '|' + 'row' + str(rx + 1) + '|上架，不能区分收货')

    return row_fcl_high, row_fcl_low, row_part_high, row_part_low, pcs_fcl_high, pcs_fcl_low, pcs_part_high, pcs_part_low, onrack_shit
# 解析拍灯表中的任务函数
def lights_workload():
    car = {}

    for rx in range(1, sh.nrows):
        name = sh.cell_value(rx, 2)
        cartoon = sh.cell_value(rx, 4)

        dict_add(name, car, cartoon)

    return car





# 读取源数据
book = xlrd.open_workbook('d:\\workload_onrack.xlsx')
sh = book.sheets()[0]
# book_person = xlrd.open_workbook('d:\\person.xlsx')
# sh_person = book_person.sheets()[0]
# person_name, person_code =[], []

# 读取人员信息
# for rx in range(sh.nrows):
#     person_name.append(str(sh.cell_value(rx, 0)))
#     person_code.append(str(sh.cell_value(rx, 1)))

# 获得表头信息
sh_header = []
for cx in range(sh.ncols):
    sh_header.append(str(sh.cell_value(0,cx)))

# 解析任务
if check_header(sh_header) == 'onrack':
    row_fcl_high, row_fcl_low, row_part_high, row_part_low, pcs_fcl_high, pcs_fcl_low, pcs_part_high, pcs_part_low, onrack_shit = onrack_workload()
    names = list(set(row_fcl_high.keys()).union(set(row_fcl_low.keys())).union(set(row_part_high.keys())).union(set(row_part_low.keys())))
elif check_header(sh_header) == 'lights':
    workload = lights_workload()

    print(workload)

    names = list(workload.keys())
else: workload = 'fuck, no load'





# 创建结果表
sheets = xlwt.Workbook( encoding="utf-8" )
sheet1 = sheets.add_sheet( "工作量统计", True )
sheet2 = sheets.add_sheet( "shit", True )

# 写入任务工作量
if check_header(sh_header) == 'onrack':
    col = 1
    set_headers = ['上架，整箱，高位(行)', '上架，整箱，高位(pcs)', '上架，整箱，低位(行)', '上架，整箱，低位(pcs)', '上架，散箱，高位(行)', '上架，散箱，高位(pcs)',
                   '上架，散箱，低位(行)', '上架，散箱，低位(pcs)']
    for header in set_headers:
        sheet1.write(0, col, header)
        col += 1

    for name in names:
        sheet1.write(names.index(name) + 1, 0, name)

        sheet1.write(names.index(name) + 1, 1, dict_val(name, row_fcl_high))
        sheet1.write(names.index(name) + 1, 2, dict_val(name, pcs_fcl_high))

        sheet1.write(names.index(name) + 1, 3, dict_val(name, row_fcl_low))
        sheet1.write(names.index(name) + 1, 4, dict_val(name, pcs_fcl_low))

        sheet1.write(names.index(name) + 1, 5, dict_val(name, row_part_high))
        sheet1.write(names.index(name) + 1, 6, dict_val(name, pcs_part_high))

        sheet1.write(names.index(name) + 1, 7, dict_val(name, row_part_low))
        sheet1.write(names.index(name) + 1, 8, dict_val(name, pcs_part_low))

    for name in onrack_shit:
        sheet2.write(onrack_shit.index(name) + 1, 0, name)
elif check_header(sh_header) == 'lights':
    col = 1
    set_headers = ['拍灯(箱)']
    for header in set_headers:
        sheet1.write(0, col, header)
        col += 1

    for name in names:
        sheet1.write(names.index(name) + 1, 0, name)
        sheet1.write(names.index(name) + 1, 1, dict_val(name, workload))
else:
    print("no load")

# 写入文件
sheets.save('d:\\stat.xls')

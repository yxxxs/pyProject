import xlrd
import pprint
import xlwt


book = xlrd.open_workbook('d:\\op.xlsx')
sh = book.sheets()[0]

sheets = xlwt.Workbook( encoding="utf-8" )
sheet1 = sheets.add_sheet("sheet1", True)

ri = 0
for cx in range(4, 17):
    for rx in range(1, 129):
        if sh.cell_value(rx, cx) == 'yes':
            sheet1.write(ri, 0, sh.cell_value(0, cx))
            sheet1.write(ri, 1, sh.cell_value(rx, 0))
            sheet1.write(ri, 2, sh.cell_value(rx, 1))
            sheet1.write(ri, 3, sh.cell_value(rx, 2))

            s = sh.cell_value(rx, 3).split('/')
            if len(s) == 1:
                sheet1.write(ri, 4, s)
                ri += 1
            else:
                sheet1.write(ri, 4, s[0])
                sheet1.write(ri+1, 4, s[1])
                ri += 2


sheets.save('d:\\po.xls')

import xlwt
import math
import pandas as pd
import numpy as np
import Data
import Functions

# 截至月（截止2017年7月各省分供应商落地份额使用率）
currentMonth = 7

#  路径前加r（原因：文件名中的 \U 开始的字符被编译器认为是八进制）
#  保存输出数据的文档地址
resultFile_path = r"C:\Users\Zhang Yu\Desktop\数据结果.xls"

# 创建结果表
wb = xlwt.Workbook(encoding='utf-8', style_compression=0)
# 创建newsheet 第二参数用于确认同一个cell单元是否可以重设值。
newsheet = wb.add_sheet('datasheet', cell_overwrite_ok=True)


#  Excel输出
#  扩大1到100列的宽度
for n in range(0, 100):
    newsheet.col(n).width = 256*11

#  表头格式
header_style = xlwt.easyxf('font: name 微软雅黑, height 220, bold on;')
#  表行列名格式
tablestyle = 'font: name 微软雅黑, height 180, bold on; '  #  粗体字
tablestyle += 'align: horz centre, vert center, wrap on; '  #  居中,自动换行
tablestyle += 'borders: left THIN, right THIN, top THIN, bottom THIN; '  #  边框
table_style = xlwt.easyxf(tablestyle)
#  正文格式
textstyle = 'font: name 微软雅黑, height 180;'  #  粗体字
textstyle += 'align: horz centre, vert center; '  #  居中
tablestyle += 'borders: left THIN, right THIN, top THIN, bottom THIN; '  #  边框
text_style = xlwt.easyxf(textstyle)


# 表头
x = 0
y = 0
newsheet.write_merge(x, x, y+0, y+99, '截止2017年%d' % currentMonth + '月各省分供应商落地份额使用率', header_style)
x += 1
newsheet.write(x, y+0, '', table_style)
newsheet.write(x, y+1, '', table_style)
newsheet.write_merge(x, x, y+2, y+4, '中标情况基础数据', table_style)
newsheet.write(x, y+5, '累计使用率', table_style)
newsheet.write(x, y+6, '', table_style)
y += 7
for n in range(0, 31):
    newsheet.write_merge(x, x, y + 3*n, y + 3*n + 2, Data.province[n], table_style)
x += 1
y = 0
newsheet.col(x).height = 256*20
newsheet.write(x, y+0, '设备类型', table_style)
newsheet.write(x, y+1, '供应商', table_style)
newsheet.write(x, y+2, '合计中标份额', table_style)
newsheet.write(x, y+3, '中标份额占比', table_style)
newsheet.write(x, y+4, '中标份额排名', table_style)
newsheet.write(x, y+5, '本期主材数量', table_style)
newsheet.write(x, y+6, '本期主材数量/C列合计中标份额*100%', table_style)
y += 7
for n in range(0, 31):
    newsheet.write(x, y+3*n, '其中：中标份额', table_style)
    newsheet.write(x, y+3*n+1, '本期主材数量', table_style)
    newsheet.write(x, y+3*n+2, '分省完成率=本省本期主材数量/本省中标份额*100%', table_style)



x += 1
y = 0
newsheet.write_merge(x, x + len(Data.dldlSupplier), y, y, Data.dldlInfo, table_style)
for n in range(0, len(Data.dldlSupplier)):
    newsheet.write(x + n, y + 1, Data.dldlSupplier[n], table_style)
    newsheet.write(x + n, y + 2, Data.dldlTotalBiddingData[n], table_style)
    newsheet.write(x + n, y + 3, xlwt.Formula('C' + str(x + n + 1) + '/C' + str(x + len(Data.dldlSupplier) + 1)),
                   table_style)
    newsheet.write(x + n, y + 4, n+1, table_style)
    newsheet.write(x + n, y + 5, xlwt.Formula(Functions.mergeStr(x + n + 1)), table_style)
    newsheet.write(x + n, y + 6, xlwt.Formula('F' + str(x + n + 1) + '/C' + str(x + len(Data.dldlSupplier) + 1)),
                   table_style)
    for m in range(0, len(Data.province)):
        newsheet.write(x + n, y + 7 + 3 * m, (0 if math.isnan(Data.dldlBiddingFrame.iloc[n, m])
                                              else Data.dldlBiddingFrame.iloc[n, m]), table_style)
    # TODO
x += len(Data.dldlSupplier)
newsheet.write(x, y + 1, '%d家合计' % len(Data.dldlSupplier), table_style)
newsheet.write(x, y + 2, Functions.totalCount(Data.dldlTotalBiddingData), table_style)
newsheet.write(x, y + 3, '100%', table_style)
newsheet.write(x, y + 4, '---', table_style)
newsheet.write(x, y + 5, xlwt.Formula(Functions.mergeStr(x + 1)), table_style)
newsheet.write(x, y + 6, xlwt.Formula('F' + str(x + 1) + '/C' + str(x + 1)), table_style)

wb.save(resultFile_path)

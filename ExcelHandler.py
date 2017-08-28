import xlwt
import Data
import Functions
import pymysql.cursors

# 创建结果表
wb = xlwt.Workbook(encoding='utf-8', style_compression=0)
# 创建newsheet 第二参数用于确认同一个cell单元是否可以重设值。
newsheet = wb.add_sheet('分省分供应商数据', cell_overwrite_ok=True)
providersheet = wb.add_sheet('6供应商数量', cell_overwrite_ok=True)
provincesheet = wb.add_sheet('9分省评价', cell_overwrite_ok=True)

# 连接数据库统计数据
connection = pymysql.connect(**Data.config)
# 使用cursor()方法获取操作游标
cursor = connection.cursor()
# 执行sql语句

# cursor.execute(Data.dldlSqlStr)
# dldlresult = cursor.fetchall()
# Functions.countByRow(dldlresult, Data.dldlFrame, Data.dldlTotalPriceFrame, Data.dldlOrderQuantityFrame,
#                      Data.dldlSupplier)
#
# cursor.execute(Data.kxSqlStr)
# kxresult = cursor.fetchall()
# Functions.countByRow(kxresult, Data.kxFrame, Data.kxTotalPriceFrame, Data.kxOrderQuantityFrame, Data.kxSupplier)
#
# cursor.execute(Data.ktSqlStr)
# ktresult = cursor.fetchall()
# Functions.countByRow(ktresult, Data.ktFrame, Data.ktTotalPriceFrame, Data.ktOrderQuantityFrame, Data.ktSupplier)
#
# cursor.execute(Data.dySqlStr)
# dyresult = cursor.fetchall()
# Functions.countByRow(dyresult, Data.dyFrame, Data.dyTotalPriceFrame, Data.dyOrderQuantityFrame, Data.dySupplier)
#
# cursor.execute(Data.wjzSqlStr)
# wjzresult = cursor.fetchall()
# Functions.countByRow(wjzresult, Data.wjzFrame, Data.wjzTotalPriceFrame, Data.wjzOrderQuantityFrame, Data.wjzSupplier)

cursor.execute(Data.ptglSqlStr)
ptglresult = cursor.fetchall()
Functions.glCountByRow(ptglresult, Data.ptglTotalPriceData, Data.ptglOrderQuanityData, Data.ptglSupplier)

cursor.execute(Data.dzglSqlStr)
dzglresult = cursor.fetchall()
Functions.glCountByRow(dzglresult, Data.dzglTotalPriceData, Data.dzglOrderQuanityData, Data.dzglSupplier)

connection.close()

# EXCEL  datasheet 输出

#  扩大1到100列的宽度
for n in range(0, 100):
    newsheet.col(n).width = 256 * 11
# 表头
x = 0
y = 0
newsheet.write_merge(x, x, y + 0, y + 99, '截止%s' % Data.end_data + '各省分供应商落地份额使用率', Data.header_style)
x += 1
newsheet.write(x, y + 0, '', Data.table_style)
newsheet.write(x, y + 1, '', Data.table_style)
newsheet.write_merge(x, x, y + 2, y + 4, '中标情况基础数据', Data.table_style)
newsheet.write(x, y + 5, '累计使用率', Data.table_style)
newsheet.write(x, y + 6, '', Data.table_style)
y += 7
for n in range(0, 31):
    newsheet.write_merge(x, x, y + 3 * n, y + 3 * n + 2, Data.province[n], Data.table_style)
x += 1
y = 0
newsheet.col(x).height = 256 * 20
newsheet.write(x, y + 0, '设备类型', Data.table_style)
newsheet.write(x, y + 1, '供应商', Data.table_style)
newsheet.write(x, y + 2, '合计中标份额', Data.table_style)
newsheet.write(x, y + 3, '中标份额占比', Data.table_style)
newsheet.write(x, y + 4, '中标份额排名', Data.table_style)
newsheet.write(x, y + 5, '本期主材数量', Data.table_style)
newsheet.write(x, y + 6, '本期主材数量/C列合计中标份额*100%', Data.table_style)
y += 7
for n in range(0, 31):
    newsheet.write(x, y + 3 * n, '其中：中标份额', Data.table_style)
    newsheet.write(x, y + 3 * n + 1, '本期主材数量', Data.table_style)
    newsheet.write(x, y + 3 * n + 2, '分省完成率=本省本期主材数量/本省中标份额*100%', Data.table_style)
x += 1
y = 0
Functions.writeByGoods(x, y, Data.dldlSupplier, Data.dldlInfo, Data.dldlTotalBiddingData, Data.dldlBiddingFrame,
                       Data.dldlFrame, newsheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.dldlSupplier) + 1)

Functions.writeByGoods(x, y, Data.kxSupplier, Data.kxInfo, Data.kxTotalBiddingData, Data.kxBiddingFrame,
                       Data.kxFrame, newsheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.kxSupplier) + 1)

Functions.writeByGoods(x, y, Data.ktSupplier, Data.ktInfo, Data.ktTotalBiddingData, Data.ktBiddingFrame,
                       Data.ktFrame, newsheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.ktSupplier) + 1)

Functions.writeByGoods(x, y, Data.dySupplier, Data.dyInfo, Data.dyTotalBiddingData, Data.dyBiddingFrame,
                       Data.dyFrame, newsheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.dySupplier) + 1)

Functions.writeByGoods(x, y, Data.wjzSupplier, Data.wjzInfo, Data.wjzTotalBiddingData, Data.wjzBiddingFrame,
                       Data.wjzFrame, newsheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.wjzSupplier) + 1)

Functions.writeByGoodsPrice(x, y, Data.ptglSupplier, Data.ptglInfo, Data.ptglTotalBiddingData, Data.ptglTotalPriceData,
                            Data.ptglTotalPrice, newsheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.ptglSupplier) + 1)

Functions.writeByGoodsPrice(x, y, Data.dzglSupplier, Data.dzglInfo, Data.dzglTotalBiddingData, Data.dzglTotalPriceData,
                            Data.dzglTotalPrice, newsheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.ptglSupplier) + 1)

# EXCEL  providersheet 输出
#  扩大1到100列的宽度
for n in range(0, 30):
    providersheet.col(n).width = 256 * 11
x = 0
y = 0
providersheet.write_merge(x, x, y + 0, y + 9, '表一：按设备类型统计', Data.table_style)
x += 1
providersheet.col(x).height = 256 * 20
providersheet.write(x, y, '序号', Data.table_style)
providersheet.write(x, y + 1, '设备类型', Data.table_style)
providersheet.write(x, y + 2, '供应商', Data.table_style)
providersheet.write(x, y + 3, '订单金额', Data.table_style)
providersheet.write(x, y + 4, '占比', Data.table_style)
providersheet.write(x, y + 5, '订单数量（指主产累计采购数量，请剔除辅材）', Data.table_style)
providersheet.write(x, y + 6, '单位', Data.table_style)
providersheet.write(x, y + 7, '占比', Data.table_style)
providersheet.write(x, y + 8, '订单笔数', Data.table_style)
providersheet.write(x, y + 9, '占比', Data.table_style)
x += 1
y = 0

a = 1
Functions.writeIn6(x, y, a, '米', Data.dldlSupplier, Data.dldlInfo, Data.dldlTotalPriceFrame, Data.dldlFrame,
                   Data.dldlOrderQuantityFrame, providersheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.dldlSupplier) + 1)
a += len(Data.dldlSupplier)

Functions.writeIn6(x, y, a, '米', Data.kxSupplier, Data.kxInfo, Data.kxTotalPriceFrame, Data.kxFrame,
                   Data.kxOrderQuantityFrame, providersheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.kxSupplier) + 1)
a += len(Data.kxSupplier)

Functions.writeIn6(x, y, a, '台', Data.ktSupplier, Data.ktInfo, Data.ktTotalPriceFrame, Data.ktFrame,
                   Data.ktOrderQuantityFrame, providersheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.ktSupplier) + 1)
a += len(Data.ktSupplier)

Functions.writeIn6(x, y, a, '套', Data.dySupplier, Data.dyInfo, Data.dyTotalPriceFrame, Data.dyFrame,
                   Data.dyOrderQuantityFrame, providersheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.dySupplier) + 1)
a += len(Data.dySupplier)

Functions.writeIn6(x, y, a, '个/套', Data.wjzSupplier, Data.wjzInfo, Data.wjzTotalPriceFrame, Data.wjzFrame,
                   Data.wjzOrderQuantityFrame, providersheet, Data.table_style, Data.text_style, Data.percent_style)
x += (len(Data.wjzSupplier) + 1)
a += len(Data.wjzSupplier)

providersheet.write_merge(x, x, y, y + 2, '总计', Data.table_style)
providersheet.write(x, y + 3, xlwt.Formula('D47+D42+D33+D22+D12'), Data.table_style)
providersheet.write(x, y + 4, '', Data.table_style)
providersheet.write(x, y + 5, xlwt.Formula('F47+F42+F33+F22+F12'), Data.table_style)
providersheet.write(x, y + 6, '', Data.table_style)
providersheet.write(x, y + 7, '', Data.table_style)
providersheet.write(x, y + 8, xlwt.Formula('I47+I42+I33+I22+I12'), Data.table_style)
providersheet.write(x, y + 9, '', Data.table_style)

x = 0
y += 12
providersheet.write_merge(x, x, y + 0, y + 4, '表二：按供应商统计', Data.table_style)
x += 1
providersheet.write_merge(x, x + 1, y, y, '序号', Data.table_style)
providersheet.write_merge(x, x + 1, y + 1, y + 1, '供应商', Data.table_style)
providersheet.write_merge(x, x, y + 2, y + 4, '截止7月31日统计结果', Data.table_style)
providersheet.write(x + 1, y + 2, '累计订单金额', Data.table_style)
providersheet.write(x + 1, y + 3, '占比', Data.table_style)
providersheet.write(x + 1, y + 4, '下单省分个数', Data.table_style)
x += 2
for n in range(0, len(Data.provider)):
    providersheet.write(x + n, y, n + 1, Data.text_style)
    providersheet.write(x + n, y + 1, Data.provider[n], Data.text_style)
    providersheet.write(x + n, y + 2, Functions.totalPriceByProvider(Data.provider[n]), Data.text_style)
    providersheet.write(x + n, y + 3, xlwt.Formula('O' + str(x + n + 1) + '/O' + str(x + 1 + len(Data.provider))),
                        Data.percent_style)
    providersheet.write(x + n, y + 4, Functions.countOrderProvince(Data.provider[n]), Data.text_style)
x += len(Data.provider)
providersheet.write(x, y, '', Data.text_style)
providersheet.write(x, y + 1, '合计', Data.text_style)
providersheet.write(x, y + 2, xlwt.Formula('SUM(O4:O36)'), Data.text_style)
providersheet.write(x, y + 3, xlwt.Formula('SUM(P4:P36)'), Data.percent_style)
providersheet.write(x, y + 4, '', Data.text_style)

# EXCEL  provincesheet 输出
for n in range(0, 30):
    provincesheet.col(n).width = 256 * 11
x = 0
y = 0
provincesheet.write(x, y, '', Data.table_style)
provincesheet.write(x, y + 1, '', Data.table_style)
provincesheet.write_merge(x, x, y + 2, y + 8, '分设备类型下单金额维度', Data.table_style)
provincesheet.write_merge(x, x, y + 9, y + 14, '分设备类型下单数量维度', Data.table_style)
provincesheet.write_merge(x, x, y + 15, y + 20, '分设备类型订单笔数维度', Data.table_style)
x += 1
provincesheet.write(x, y + 0, '序号', Data.table_style)
provincesheet.write(x, y + 1, '单位', Data.table_style)
provincesheet.write(x, y + 2, '电缆（万元）', Data.table_style)
provincesheet.write(x, y + 3, '电源（万元）', Data.table_style)
provincesheet.write(x, y + 4, '空调（万元）', Data.table_style)
provincesheet.write(x, y + 5, '馈线（万元）', Data.table_style)
provincesheet.write(x, y + 6, '微基站pRRU（万元）', Data.table_style)
provincesheet.write(x, y + 7, '总计（万元）', Data.table_style)
provincesheet.write(x, y + 8, '占比', Data.table_style)
provincesheet.write(x, y + 9, '电缆（米）', Data.table_style)
provincesheet.write(x, y + 10, '电源（套）', Data.table_style)
provincesheet.write(x, y + 11, '空调（台）', Data.table_style)
provincesheet.write(x, y + 12, '馈线（千米）', Data.table_style)
provincesheet.write(x, y + 13, '微基站pRRU（个/套）', Data.table_style)
provincesheet.write(x, y + 14, '总计（项）', Data.table_style)
provincesheet.write(x, y + 15, '电缆（笔）', Data.table_style)
provincesheet.write(x, y + 16, '电源（笔）', Data.table_style)
provincesheet.write(x, y + 17, '空调（笔）', Data.table_style)
provincesheet.write(x, y + 18, '馈线（笔）', Data.table_style)
provincesheet.write(x, y + 19, '微基站pRRU（笔）', Data.table_style)
provincesheet.write(x, y + 20, '总计（笔）', Data.table_style)
x += 1
for n in range(0, len(Data.province)):
    Functions.writeIn9(x, y, n, provincesheet)
    x += 1
provincesheet.write_merge(x, x, y, y + 1, '合计', Data.table_style)
provincesheet.write(x, y + 2, xlwt.Formula('SUM(C3:C33)'), Data.text_style)
provincesheet.write(x, y + 3, xlwt.Formula('SUM(D3:D33)'), Data.text_style)
provincesheet.write(x, y + 4, xlwt.Formula('SUM(E3:E33)'), Data.text_style)
provincesheet.write(x, y + 5, xlwt.Formula('SUM(F3:F33)'), Data.text_style)
provincesheet.write(x, y + 6, xlwt.Formula('SUM(G3:G33)'), Data.text_style)
provincesheet.write(x, y + 7, xlwt.Formula('SUM(H3:H33)'), Data.text_style)
provincesheet.write(x, y + 8, xlwt.Formula('SUM(I3:I33)'), Data.percent_style)
provincesheet.write(x, y + 9, xlwt.Formula('SUM(J3:J33)'), Data.text_style)
provincesheet.write(x, y + 10, xlwt.Formula('SUM(K3:K33)'), Data.text_style)
provincesheet.write(x, y + 11, xlwt.Formula('SUM(L3:L33)'), Data.text_style)
provincesheet.write(x, y + 12, xlwt.Formula('SUM(M3:M33)'), Data.text_style)
provincesheet.write(x, y + 13, xlwt.Formula('SUM(N3:N33)'), Data.text_style)
provincesheet.write(x, y + 14, xlwt.Formula('SUM(O3:O33)'), Data.text_style)
provincesheet.write(x, y + 15, xlwt.Formula('SUM(P3:P33)'), Data.text_style)
provincesheet.write(x, y + 16, xlwt.Formula('SUM(Q3:Q33)'), Data.text_style)
provincesheet.write(x, y + 17, xlwt.Formula('SUM(R3:R33)'), Data.text_style)
provincesheet.write(x, y + 18, xlwt.Formula('SUM(S3:S33)'), Data.text_style)
provincesheet.write(x, y + 19, xlwt.Formula('SUM(T3:T33)'), Data.text_style)
provincesheet.write(x, y + 20, xlwt.Formula('SUM(U3:U33)'), Data.text_style)

wb.save(Data.resultFile_path)

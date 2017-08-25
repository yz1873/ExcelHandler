import xlwt
import Data
import Functions
import pymysql.cursors

# 创建结果表
wb = xlwt.Workbook(encoding='utf-8', style_compression=0)
# 创建newsheet 第二参数用于确认同一个cell单元是否可以重设值。
newsheet = wb.add_sheet('datasheet', cell_overwrite_ok=True)

# 连接数据库统计数据
connection = pymysql.connect(**Data.config)
# 使用cursor()方法获取操作游标
cursor = connection.cursor()
# 执行sql语句
try:
    # 执行sql语句，进行查询

    cursor.execute(Data.dldlSqlStr)
    dldlresult = cursor.fetchall()
    Functions.countByRow(dldlresult, Data.dldlFrame, Data.dldlSupplier)

    # cursor.execute(Data.kxSqlStr)
    # kxresult = cursor.fetchall()
    # Functions.countByRow(kxresult, Data.kxFrame, Data.kxSupplier)

    # cursor.execute(Data.ktSqlStr)
    # ktresult = cursor.fetchall()
    # Functions.countByRow(ktresult, Data.ktFrame, Data.ktSupplier)

    # cursor.execute(Data.dySqlStr)
    # dyresult = cursor.fetchall()
    # Functions.countByRow(dyresult, Data.dyFrame, Data.dySupplier)
    # print(dyresult)
    # print(Data.dyFrame)

except:
    print("Error: unable to fecth data")
connection.close()

#  Excel输出
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

wb.save(Data.resultFile_path)

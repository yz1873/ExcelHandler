import math
import Data
import xlwt

def totalCount(list):
    result = 0
    for n in range(0, len(list)):
        result += list[n]
    return result

def intTran(dataframeloc):
    return 0 if math.isnan(dataframeloc) else int(dataframeloc)

def totalCountByProvince(m, dataframe):
    return (0 if math.isnan(dataframe[Data.province[m]].T.sum())
            else int(dataframe[Data.province[m]].T.sum()))

def mergeStr(n):
    return 'I'+str(n)+'+L'+str(n)+'+O'+str(n)+'+R'+str(n)+'+U'+str(n)+'+X'+str(n)+'+AA'+str(n)+'+AD'+str(n)+'+AG'+str(n)+'+AJ'+str(n)+'+AM'\
           +str(n)+'+AP'+str(n)+'+AS'+str(n)+'+AV'+str(n)+'+AY'+str(n)+'+BB'+str(n)+'+BE'+str(n)+'+BH'+str(n)+'+BK'+str(n)+'+BN'+str(n)+'+BQ'\
           +str(n)+'+BT'+str(n)+'+BW'+str(n)+'+BZ'+str(n)+'+CC'+str(n)+'+CF'+str(n)+'+CI'+str(n)+'+CL'+str(n)+'+CO'+str(n)+'+CR'+str(n)+'+CU'+str(n)

def writeByGoods(x, y, GoodsSupplier, GoodsInfo, GoodsTotalBiddingData, GoodsBiddingFrame, GoodsFrame, currentsheet,
                 table_style, text_style, percent_style):
    currentsheet.write_merge(x, x + len(GoodsSupplier), y, y, GoodsInfo, table_style)
    for n in range(0, len(GoodsSupplier)):
        currentsheet.write(x + n, y + 1, GoodsSupplier[n], table_style)
        currentsheet.write(x + n, y + 2, GoodsTotalBiddingData[n], text_style)
        currentsheet.write(x + n, y + 3, xlwt.Formula('C' + str(x + n + 1) + '/C' + str(x + len(GoodsSupplier) + 1)),
                       percent_style)
        currentsheet.write(x + n, y + 4, n+1, text_style)
        currentsheet.write(x + n, y + 5, xlwt.Formula(mergeStr(x + n + 1)), text_style)
        currentsheet.write(x + n, y + 6, xlwt.Formula('F' + str(x + n + 1) + '/C' + str(x + len(GoodsSupplier) + 1)),
                       percent_style)
        for m in range(0, len(Data.province)):
            currentsheet.write(x + n, y + 7 + 3 * m, (intTran(GoodsBiddingFrame.iloc[n, m])), text_style)
            currentsheet.write(x + n, y + 8 + 3 * m, (intTran(GoodsFrame.iloc[n, m])), text_style)
            currentsheet.write(x + n, y + 9 + 3 * m, (0 if (intTran(GoodsBiddingFrame.iloc[n, m]) == 0)
                                                  else (intTran(GoodsFrame.iloc[n, m])/
                                                        intTran(GoodsBiddingFrame.iloc[n, m]))), percent_style)
    x += len(GoodsSupplier)
    currentsheet.write(x, y + 1, '%d家合计' % len(GoodsSupplier), table_style)
    currentsheet.write(x, y + 2, totalCount(GoodsTotalBiddingData), text_style)
    currentsheet.write(x, y + 3, 1, percent_style)
    currentsheet.write(x, y + 4, '---', table_style)
    currentsheet.write(x, y + 5, xlwt.Formula(mergeStr(x + 1)), text_style)
    currentsheet.write(x, y + 6, xlwt.Formula('F' + str(x + 1) + '/C' + str(x + 1)), percent_style)
    for m in range(0, len(Data.province)):
        currentsheet.write(x, y + 7 + 3 * m, totalCountByProvince(m, GoodsBiddingFrame), text_style)
        currentsheet.write(x, y + 8 + 3 * m, totalCountByProvince(m, GoodsFrame), text_style)
        currentsheet.write(x, y + 9 + 3 * m, (0 if (totalCountByProvince(m, GoodsBiddingFrame) == 0)
                                          else (totalCountByProvince(m, GoodsFrame)/
                                                totalCountByProvince(m, GoodsBiddingFrame))), percent_style)

def countByRow(result, dataframe, supplier):    # TODO 永不用判断是NaN
    for row in result:
        if math.isnan(dataframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])]):
            dataframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])] = row['采购数量']
        else:
            dataframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])] += row['采购数量']
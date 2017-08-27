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


def totalCountByProvider(m, providerlist, dataframe):
    return (0 if math.isnan(dataframe.loc[providerlist[m]].sum())
            else int(dataframe.loc[providerlist[m]].sum()))


def mergeStr(n):
    return 'I' + str(n) + '+L' + str(n) + '+O' + str(n) + '+R' + str(n) + '+U' + str(n) + '+X' + str(n) + '+AA' + str(
        n) + '+AD' + str(n) + '+AG' + str(n) + '+AJ' + str(n) + '+AM' \
           + str(n) + '+AP' + str(n) + '+AS' + str(n) + '+AV' + str(n) + '+AY' + str(n) + '+BB' + str(n) + '+BE' + str(
        n) + '+BH' + str(n) + '+BK' + str(n) + '+BN' + str(n) + '+BQ' \
           + str(n) + '+BT' + str(n) + '+BW' + str(n) + '+BZ' + str(n) + '+CC' + str(n) + '+CF' + str(n) + '+CI' + str(
        n) + '+CL' + str(n) + '+CO' + str(n) + '+CR' + str(n) + '+CU' + str(n)


def writeByGoods(x, y, GoodsSupplier, GoodsInfo, GoodsTotalBiddingData, GoodsBiddingFrame, GoodsFrame, currentsheet,
                 table_style, text_style, percent_style):
    currentsheet.write_merge(x, x + len(GoodsSupplier), y, y, GoodsInfo, table_style)
    for n in range(0, len(GoodsSupplier)):
        currentsheet.write(x + n, y + 1, GoodsSupplier[n], table_style)
        currentsheet.write(x + n, y + 2, GoodsTotalBiddingData[n], text_style)
        currentsheet.write(x + n, y + 3, xlwt.Formula('C' + str(x + n + 1) + '/C' + str(x + len(GoodsSupplier) + 1)),
                           percent_style)
        currentsheet.write(x + n, y + 4, n + 1, text_style)
        currentsheet.write(x + n, y + 5, xlwt.Formula(mergeStr(x + n + 1)), text_style)
        currentsheet.write(x + n, y + 6, xlwt.Formula('F' + str(x + n + 1) + '/C' + str(x + len(GoodsSupplier) + 1)),
                           percent_style)
        for m in range(0, len(Data.province)):
            currentsheet.write(x + n, y + 7 + 3 * m, (intTran(GoodsBiddingFrame.iloc[n, m])), text_style)
            currentsheet.write(x + n, y + 8 + 3 * m, (intTran(GoodsFrame.iloc[n, m])), text_style)
            currentsheet.write(x + n, y + 9 + 3 * m, (0 if (intTran(GoodsBiddingFrame.iloc[n, m]) == 0)
                                                      else (intTran(GoodsFrame.iloc[n, m]) /
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
                                              else (totalCountByProvince(m, GoodsFrame) /
                                                    totalCountByProvince(m, GoodsBiddingFrame))), percent_style)


def writeByGoodsPrice(x, y, GoodsSupplier, GoodsInfo, GoodsTotalBiddingData, GoodsSellData, currentsheet, table_style,
                      text_style, percent_style):
    currentsheet.write_merge(x, x + len(GoodsSupplier), y, y, GoodsInfo, table_style)
    for n in range(0, len(GoodsSupplier)):
        currentsheet.write(x + n, y + 1, GoodsSupplier[n], table_style)
        currentsheet.write(x + n, y + 2, GoodsTotalBiddingData[n], text_style)
        currentsheet.write(x + n, y + 3, xlwt.Formula('C' + str(x + n + 1) + '/C' + str(x + len(GoodsSupplier) + 1)),
                           percent_style)
        currentsheet.write(x + n, y + 4, n + 1, text_style)
        currentsheet.write(x + n, y + 5, GoodsSellData[n], text_style)
        currentsheet.write(x + n, y + 6, xlwt.Formula('F' + str(x + n + 1) + '/C' + str(x + len(GoodsSupplier) + 1)),
                           percent_style)
    x += len(GoodsSupplier)
    currentsheet.write(x, y + 1, '%d家合计' % len(GoodsSupplier), table_style)
    currentsheet.write(x, y + 2, totalCount(GoodsTotalBiddingData), text_style)
    currentsheet.write(x, y + 3, 1, percent_style)
    currentsheet.write(x, y + 4, '---', table_style)
    currentsheet.write(x, y + 5, totalCount(GoodsSellData), text_style)
    currentsheet.write(x, y + 6, xlwt.Formula('F' + str(x + 1) + '/C' + str(x + 1)), percent_style)


def countByRow(result, dataframe, totalpriceframe, quantityframe, supplier):
    for row in result:
        if math.isnan(dataframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])]):
            dataframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])] = row['采购数量']
        else:
            dataframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])] += row['采购数量']

        if math.isnan(totalpriceframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])]):
            totalpriceframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])] = row['价税合计']
        else:
            totalpriceframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])] += row['价税合计']

        if math.isnan(quantityframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])]):
            quantityframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])] = 1
        else:
            quantityframe.iloc[supplier.index(row['供应商']), Data.province.index(row['省分公司'])] += 1


def writeIn6(x, y, a, unit, GoodsSupplier, GoodsInfo, totalpriceframe, goodsframe, orderquantityframe, currentsheet,
             table_style, text_style, percent_style):
    for n in range(0, len(GoodsSupplier)):
        currentsheet.write(x + n, y, n + a, table_style)
        currentsheet.write(x + n, y + 2, GoodsSupplier[n], table_style)
        currentsheet.write(x + n, y + 3, totalCountByProvider(n, GoodsSupplier, totalpriceframe), text_style)
        currentsheet.write(x + n, y + 4, xlwt.Formula('D' + str(x + n + 1) + '/D' + str(x + len(GoodsSupplier) + 1)),
                           percent_style)
        currentsheet.write(x + n, y + 5, totalCountByProvider(n, GoodsSupplier, goodsframe), text_style)
        currentsheet.write(x + n, y + 6, unit, text_style)
        currentsheet.write(x + n, y + 7, xlwt.Formula('F' + str(x + n + 1) + '/F' + str(x + len(GoodsSupplier) + 1)),
                           percent_style)
        currentsheet.write(x + n, y + 8, totalCountByProvider(n, GoodsSupplier, orderquantityframe), text_style)
        currentsheet.write(x + n, y + 9, xlwt.Formula('I' + str(x + n + 1) + '/I' + str(x + len(GoodsSupplier) + 1)),
                           percent_style)
    currentsheet.write(x + len(GoodsSupplier), y, '', table_style)
    currentsheet.write_merge(x, x + len(GoodsSupplier), y + 1, y + 1, GoodsInfo, table_style)
    currentsheet.write(x + len(GoodsSupplier), y + 2, '%d家合计' % len(GoodsSupplier), table_style)
    currentsheet.write(x + len(GoodsSupplier), y + 3,
                       xlwt.Formula('SUM(D' + str(x + 1) + ':D' + str(x + len(GoodsSupplier)) + ')'), text_style)
    currentsheet.write(x + len(GoodsSupplier), y + 4,
                       xlwt.Formula('SUM(E' + str(x + 1) + ':E' + str(x + len(GoodsSupplier)) + ')'), percent_style)
    currentsheet.write(x + len(GoodsSupplier), y + 5,
                       xlwt.Formula('SUM(F' + str(x + 1) + ':F' + str(x + len(GoodsSupplier)) + ')'), text_style)
    currentsheet.write(x + len(GoodsSupplier), y + 6, unit, text_style)
    currentsheet.write(x + len(GoodsSupplier), y + 7,
                       xlwt.Formula('SUM(H' + str(x + 1) + ':H' + str(x + len(GoodsSupplier)) + ')'), percent_style)
    currentsheet.write(x + len(GoodsSupplier), y + 8,
                       xlwt.Formula('SUM(I' + str(x + 1) + ':I' + str(x + len(GoodsSupplier)) + ')'), text_style)
    currentsheet.write(x + len(GoodsSupplier), y + 9,
                       xlwt.Formula('SUM(J' + str(x + 1) + ':J' + str(x + len(GoodsSupplier)) + ')'), percent_style)

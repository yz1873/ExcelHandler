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


def writeByGoodsPrice(x, y, GoodsSupplier, GoodsInfo, GoodsTotalBiddingData, GoodsSellData, totalCountData,
                      currentsheet, table_style, text_style, percent_style):
    currentsheet.write_merge(x, x + len(GoodsSupplier), y, y, GoodsInfo, table_style)
    for n in range(0, len(GoodsSupplier)):
        currentsheet.write(x + n, y + 1, GoodsSupplier[n], table_style)
        currentsheet.write(x + n, y + 2, xlwt.Formula('D' + str(x + n + 1) + '*C' + str(x + len(GoodsSupplier) + 1)),
                           text_style)
        currentsheet.write(x + n, y + 3, GoodsTotalBiddingData[n], percent_style)
        currentsheet.write(x + n, y + 4, '', text_style)
        currentsheet.write(x + n, y + 5, GoodsSellData[n], text_style)
        currentsheet.write(x + n, y + 6, xlwt.Formula('F' + str(x + n + 1) + '/C' + str(x + len(GoodsSupplier) + 1)),
                           percent_style)
    x += len(GoodsSupplier)
    currentsheet.write(x, y + 1, '%d家合计' % len(GoodsSupplier), table_style)
    currentsheet.write(x, y + 2, totalCountData, text_style)
    currentsheet.write(x, y + 3, 1, percent_style)
    currentsheet.write(x, y + 4, '', table_style)
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


def glCountByRow(result, pricedata, quantitydata, supplier):
    for row in result:
        pricedata[supplier.index(row['供应商'])] = row['价税合计']
        quantitydata[supplier.index(row['供应商'])] = row['订单数']


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


def totalPriceByProvider(providerName):
    result = 0
    if providerName in Data.dldlTotalPriceFrame.index:
        result += (0 if math.isnan(Data.dldlTotalPriceFrame.loc[providerName].sum()) else int(
            Data.dldlTotalPriceFrame.loc[providerName].sum()))
    if providerName in Data.kxTotalPriceFrame.index:
        result += (0 if math.isnan(Data.kxTotalPriceFrame.loc[providerName].sum()) else int(
            Data.kxTotalPriceFrame.loc[providerName].sum()))
    if providerName in Data.ktTotalPriceFrame.index:
        result += (0 if math.isnan(Data.ktTotalPriceFrame.loc[providerName].sum()) else int(
            Data.ktTotalPriceFrame.loc[providerName].sum()))
    if providerName in Data.dyTotalPriceFrame.index:
        result += (0 if math.isnan(Data.dyTotalPriceFrame.loc[providerName].sum()) else int(
            Data.dyTotalPriceFrame.loc[providerName].sum()))
    if providerName in Data.wjzTotalPriceFrame.index:
        result += (0 if math.isnan(Data.wjzTotalPriceFrame.loc[providerName].sum()) else int(
            Data.wjzTotalPriceFrame.loc[providerName].sum()))
    return result


def countOrderProvince(providerName):
    result = set()
    if providerName in Data.dldlTotalPriceFrame.index:
        for m in range(0, len(Data.province)):
            if (intTran(Data.dldlTotalPriceFrame.loc[providerName, Data.province[m]]) != 0):
                result.add(Data.province[m])
    if providerName in Data.kxTotalPriceFrame.index:
        for m in range(0, len(Data.province)):
            if (intTran(Data.kxTotalPriceFrame.loc[providerName, Data.province[m]]) != 0):
                result.add(Data.province[m])
    if providerName in Data.ktTotalPriceFrame.index:
        for m in range(0, len(Data.province)):
            if (intTran(Data.ktTotalPriceFrame.loc[providerName, Data.province[m]]) != 0):
                result.add(Data.province[m])
    if providerName in Data.dyTotalPriceFrame.index:
        for m in range(0, len(Data.province)):
            if (intTran(Data.dyTotalPriceFrame.loc[providerName, Data.province[m]]) != 0):
                result.add(Data.province[m])
    if providerName in Data.wjzTotalPriceFrame.index:
        for m in range(0, len(Data.province)):
            if (intTran(Data.wjzTotalPriceFrame.loc[providerName, Data.province[m]]) != 0):
                result.add(Data.province[m])
    return len(result)


def writeIn9(x, y, n, currentsheet):
    currentsheet.write(x, y, n + 1, Data.text_style)
    currentsheet.write(x, y + 1, Data.province[n], Data.text_style)
    currentsheet.write(x, y + 2, round(totalCountByProvince(n, Data.dldlTotalPriceFrame) / 10000, 2), Data.text_style)
    currentsheet.write(x, y + 3, round(totalCountByProvince(n, Data.dyTotalPriceFrame) / 10000, 2), Data.text_style)
    currentsheet.write(x, y + 4, round(totalCountByProvince(n, Data.ktTotalPriceFrame) / 10000, 2), Data.text_style)
    currentsheet.write(x, y + 5, round(totalCountByProvince(n, Data.kxTotalPriceFrame) / 10000, 2), Data.text_style)
    currentsheet.write(x, y + 6, round(totalCountByProvince(n, Data.wjzTotalPriceFrame) / 10000, 2), Data.text_style)
    currentsheet.write(x, y + 7, xlwt.Formula(
        'C' + str(x + 1) + '+D' + str(x + 1) + '+E' + str(x + 1) + '+F' + str(x + 1) + '+G' + str(x + 1)),
                       Data.text_style)
    currentsheet.write(x, y + 8, xlwt.Formula('H' + str(x + 1) + '/H34'), Data.percent_style)
    currentsheet.write(x, y + 9, totalCountByProvince(n, Data.dldlFrame), Data.text_style)
    currentsheet.write(x, y + 10, totalCountByProvince(n, Data.dyFrame), Data.text_style)
    currentsheet.write(x, y + 11, totalCountByProvince(n, Data.ktFrame), Data.text_style)
    currentsheet.write(x, y + 12, round(totalCountByProvince(n, Data.kxFrame) / 1000, 2), Data.text_style)
    currentsheet.write(x, y + 13, totalCountByProvince(n, Data.wjzFrame), Data.text_style)
    currentsheet.write(x, y + 14, xlwt.Formula(
        'J' + str(x + 1) + '+K' + str(x + 1) + '+L' + str(x + 1) + '+M' + str(x + 1) + '+N' + str(x + 1)),
                       Data.text_style)
    currentsheet.write(x, y + 15, totalCountByProvince(n, Data.dldlOrderQuantityFrame), Data.text_style)
    currentsheet.write(x, y + 16, totalCountByProvince(n, Data.dyOrderQuantityFrame), Data.text_style)
    currentsheet.write(x, y + 17, totalCountByProvince(n, Data.ktOrderQuantityFrame), Data.text_style)
    currentsheet.write(x, y + 18, totalCountByProvince(n, Data.kxOrderQuantityFrame), Data.text_style)
    currentsheet.write(x, y + 19, totalCountByProvince(n, Data.wjzOrderQuantityFrame), Data.text_style)
    currentsheet.write(x, y + 20, xlwt.Formula(
        'P' + str(x + 1) + '+Q' + str(x + 1) + '+R' + str(x + 1) + '+S' + str(x + 1) + '+T' + str(x + 1)),
                       Data.text_style)

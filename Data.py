import pandas as pd
import xlwt
import pymysql.cursors
import datetime

# 截至日
# end_data = datetime.date(2017, 7, 31)
end_data = datetime.datetime(2017, 7, 31, 23, 59, 59)
#  路径前加r（原因：文件名中的 \U 开始的字符被编译器认为是八进制）
#  保存输出数据的文档地址  Administrator
resultFile_path = r"C:\Users\Administrator\Desktop\数据结果.xls"
# resultFile_path = r"C:\Users\Zhang Yu\Desktop\数据结果.xls"

config = {
    'host': '10.0.204.205',
    'port': 3306,
    'user': 'read2Pan20170217',
    'password': 'Read2Pan20170217',
    'db': 'eshop',
    'charset': 'utf8mb4',
    'cursorclass': pymysql.cursors.DictCursor,
}

# 所有省分
province = ['北京', '天津', '河北', '山西', '内蒙', '辽宁', '吉林', '黑龙', '上海', '江苏', '浙江',
            '安徽', '福建', '江西', '山东', '河南', '湖北', '湖南', '广东', '广西', '海南', '重庆',
            '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海', '宁夏', '新疆']

# 电力电缆（简称dldl）
dldlSqlStr = "select LEFT(o.province_name,2) '省分公司', p.real_nums '采购数量'," \
             "round(round(p.real_nums * p.company_price,5),2) AS '价税合计', " \
             "CASE " \
             "WHEN p.provider_id = 49230 THEN '江苏中利' " \
             "WHEN p.provider_id = 49256 THEN '通鼎互联' " \
             "WHEN p.provider_id = 49214 THEN '江苏俊知' " \
             "WHEN p.provider_id = 49227 THEN '中天科技' " \
             "WHEN p.provider_id = 49224 THEN '江苏亨通' " \
             "WHEN p.provider_id = 49228 THEN '西部电缆' " \
             "WHEN p.provider_id = 49231 THEN '鲁能泰山' " \
             "WHEN p.provider_id = 49210 THEN '成都大唐' " \
             "WHEN p.provider_id = 49226 THEN '富通集团' " \
             "END '供应商' " \
             "from eshop_order_product p " \
             "LEFT JOIN eshop_order o ON p.order_id = o.id " \
             "LEFT JOIN eshop_provideraddress epa ON epa.providerId = p.provider_id " \
             "LEFT JOIN eshop_provider_contact c ON c.provider_id = p.provider_id " \
             "LEFT JOIN eshop_goods g ON g.item_all = p.ITEM_NUMBER " \
             "WHERE epa.shop_id = o.shop_id " \
             "AND c.shop_id = o.shop_id " \
             "AND g.shop_id = o.shop_id " \
             "AND o.shop_id = ' 596 ' " \
             "AND p.CONTACT_NUMBER = c.contact_number " \
             "AND p.CONTACT_NUMBER IN ('CU12-1001-2016-001073','CU12-1001-2016-001074'," \
             "'CU12-1001-2016-001075','CU12-1001-2016-001076','CU12-1001-2016-001077'," \
             "'CU12-1001-2016-001078','CU12-1001-2016-001079','CU12-1001-2016-001080','CU12-1001-2016-001081') " \
             "AND o.`status` in ('2','5') " \
             "AND o.create_time BETWEEN '2016-01-01' And '%s' " \
             "GROUP BY p.id " % end_data

dldlInfo = "电力电缆（单位：米）  备注：江苏中利集团股份有限即中利科技"
dldlSupplier = ['江苏亨通', '中天科技', '江苏俊知', '成都大唐', '富通集团', '通鼎互联', '江苏中利', '鲁能泰山', '西部电缆']

dldlTotalBiddingData = [2249329, 2164380, 2156859, 1102940, 1082391, 827624, 803870, 181140, 176960]

dldlBiddingData = {'北京': [17151, 16500, 16443, 0, 0, 15881, 15425, 0, 0],
                   '天津': [373360, 359184, 357944, 0, 0, 345717, 335794, 0, 0],
                   '河北': [58204, 55994, 55800, 0, 0, 53894, 52347, 0, 0],
                   '山西': [198901, 191349, 190688, 0, 0, 184174, 178888, 0, 0],
                   '内蒙': [71311, 68604, 68367, 0, 0, 66032, 64136, 0, 0],
                   '辽宁': [55571, 53482, 53295, 53161, 52171, 0, 0, 0, 0],
                   '吉林': [24165, 23257, 23175, 23117, 22686, 0, 0, 0, 0],
                   '黑龙': [22686, 21834, 21758, 21703, 21299, 0, 0, 0, 0],
                   '上海': [147983, 142423, 141924, 141568, 138930, 0, 0, 0, 0],
                   '江苏': [16525, 15904, 15848, 15809, 15514, 0, 0, 0, 0],
                   '浙江': [216850, 208703, 207972, 207450, 203585, 0, 0, 0, 0],
                   '安徽': [38391, 36949, 36819, 36727, 36042, 0, 0, 0, 0],
                   '福建': [21591, 20779, 20706, 20654, 20270, 0, 0, 0, 0],
                   '江西': [3508, 3377, 3365, 3356, 3294, 0, 0, 0, 0],
                   '山东': [432123, 415888, 414431, 413390, 405688, 0, 0, 0, 0],
                   '河南': [64111, 61677, 61467, 0, 0, 0, 0, 57319, 55996],
                   '湖北': [17428, 16767, 16710, 0, 0, 0, 0, 15582, 15223],
                   '湖南': [67070, 64523, 64303, 0, 0, 0, 0, 59964, 58580],
                   '广东': [26770, 25754, 25666, 0, 0, 0, 0, 23934, 23382],
                   '广西': [13651, 13133, 13088, 0, 0, 0, 0, 12205, 11923],
                   '海南': [13574, 13058, 13014, 0, 0, 0, 0, 12135, 11855],
                   '重庆': [26127, 25135, 25048, 0, 0, 24192, 23498, 0, 0],
                   '四川': [33712, 32432, 32320, 0, 0, 31216, 30320, 0, 0],
                   '贵州': [32462, 31229, 31121, 0, 0, 30058, 29196, 0, 0],
                   '云南': [45828, 44087, 43935, 0, 0, 42434, 41216, 0, 0],
                   '西藏': [36746, 35351, 35229, 0, 0, 34025, 33049, 0, 0],
                   '陕西': [124145, 119480, 119062, 118763, 116550, 0, 0, 0, 0],
                   '甘肃': [8852, 8519, 8490, 8468, 8311, 0, 0, 0, 0],
                   '青海': [5238, 5041, 5023, 5011, 4917, 0, 0, 0, 0],
                   '宁夏': [6768, 6513, 6491, 6474, 6354, 0, 0, 0, 0],
                   '新疆': [28524, 27453, 27356, 27288, 26779, 0, 0, 0, 0]}
dldlBiddingFrame = pd.DataFrame(dldlBiddingData, columns=province, index=dldlSupplier)

dldlData = {}
dldlFrame = pd.DataFrame(dldlData, columns=province, index=dldlSupplier)

dldlTotalPrice = {}
dldlTotalPriceFrame = pd.DataFrame(dldlTotalPrice, columns=province, index=dldlSupplier)

dldlOrderQuantity = {}
dldlOrderQuantityFrame = pd.DataFrame(dldlOrderQuantity, columns=province, index=dldlSupplier)

# 馈线（简称kx）
kxSqlStr = "select LEFT(o.province_name,2) '省分公司', p.real_nums '采购数量', " \
           "round(round(p.real_nums * p.company_price,5),2) AS '价税合计', " \
           "LEFT(p.provider_name,4) '供应商'  " \
           "from eshop_order_product p " \
           "LEFT JOIN eshop_order o ON p.order_id = o.id " \
           "LEFT JOIN eshop_provideraddress epa ON epa.providerId = p.provider_id " \
           "LEFT JOIN eshop_provider_contact c ON c.provider_id = p.provider_id " \
           "LEFT JOIN eshop_goods g ON g.item_all = p.ITEM_NUMBER " \
           "LEFT JOIN eshop_materials_catergorytree mac ON mac.id =p.goodstype_id " \
           "WHERE epa.shop_id = o.shop_id " \
           "AND c.shop_id = o.shop_id " \
           "AND g.shop_id = o.shop_id " \
           "AND o.shop_id = ' 596 ' " \
           "AND p.CONTACT_NUMBER = c.contact_number " \
           "AND p.CONTACT_NUMBER IN ('CU12-1001-2016-001103','CU12-1001-2016-001104','CU12-1001-2016-001105'," \
           "'CU12-1001-2016-001106','CU12-1001-2016-001107','CU12-1001-2016-001108'," \
           "'CU12-1001-2016-001109','CU12-1001-2016-001110','CU12-1001-2016-001111')" \
           "AND o.`status` in ('2','5') " \
           "AND mac.name = '铜缆' " \
           "AND o.create_time BETWEEN '2016-01-01' And '%s' " \
           "GROUP BY p.id " % end_data
kxInfo = "馈线（单位：米）"
kxSupplier = ['江苏俊知', '江苏亨鑫', '珠海汉胜', '长飞光纤', '通鼎互联',
              '中天射频', '成都大唐', '富通集团', '湖北凯乐']

kxTotalBiddingData = [41599100, 41117902, 39462841, 21376628, 21081553, 10225762, 10036104, 7042699, 7016808]

kxBiddingData = {'北京': [1439702, 1423107, 1365713, 1352574, 1333904, 0, 0, 0, 0],
                 '天津': [1723480, 1703612, 1634905, 1619177, 1596826, 0, 0, 0, 0],
                 '河北': [2413259, 2385440, 2289234, 2267211, 2235915, 0, 0, 0, 0],
                 '山西': [98478, 97343, 93418, 92519, 91242, 0, 0, 0, 0],
                 '内蒙': [500720, 494949, 474988, 470418, 463925, 0, 0, 0, 0],
                 '辽宁': [3129804, 3092739, 2968199, 0, 0, 0, 0, 2822902, 2812524],
                 '吉林': [746450, 737610, 707907, 0, 0, 0, 0, 673254, 670779],
                 '黑龙': [1241805, 1227100, 1177687, 0, 0, 0, 0, 1120038, 1115920],
                 '上海': [984682, 973331, 934076, 925090, 912321, 0, 0, 0, 0],
                 '江苏': [5545161, 5481239, 5260179, 5209574, 5137663, 0, 0, 0, 0],
                 '浙江': [2021686, 1998382, 1917786, 1899336, 1873119, 0, 0, 0, 0],
                 '安徽': [1999957, 1976903, 1897173, 1878922, 1852986, 0, 0, 0, 0],
                 '福建': [1517719, 1500224, 1439719, 1425869, 1406186, 0, 0, 0, 0],
                 '江西': [145740, 144060, 138250, 136920, 135030, 0, 0, 0, 0],
                 '山东': [2370773, 2343445, 2248933, 2227297, 2196552, 0, 0, 0, 0],
                 '河南': [1257000, 1242600, 1192800, 0, 0, 1164600, 1143000, 0, 0],
                 '湖北': [1835701, 1814671, 1741944, 0, 0, 1700761, 1669217, 0, 0],
                 '湖南': [407269, 402602, 386467, 0, 0, 377330, 370332, 0, 0],
                 '广东': [6892550, 6813590, 6540520, 0, 0, 6385890, 6267450, 0, 0],
                 '广西': [518863, 512918, 492362, 0, 0, 480721, 471805, 0, 0],
                 '海南': [125700, 124260, 119280, 0, 0, 116460, 114300, 0, 0],
                 '重庆': [527995, 521909, 500860, 496042, 489194, 0, 0, 0, 0],
                 '四川': [81366, 80428, 77185, 76442, 75387, 0, 0, 0, 0],
                 '贵州': [1287149, 1272312, 1220999, 1209253, 1192561, 0, 0, 0, 0],
                 '云南': [86245, 85251, 81813, 81026, 79907, 0, 0, 0, 0],
                 '西藏': [9535, 9426, 9046, 8958, 8835, 34025, 33049, 0, 0],
                 '陕西': [2019456, 1995541, 1915183, 0, 0, 0, 0, 1821433, 1814737],
                 '甘肃': [11610, 11473, 11011, 0, 0, 0, 0, 10472, 10434],
                 '青海': [46653, 46101, 44244, 0, 0, 0, 0, 42078, 41924],
                 '宁夏': [33776, 33376, 32032, 0, 0, 0, 0, 30464, 30352],
                 '新疆': [578816, 571960, 548928, 0, 0, 0, 0, 522058, 520138]}
kxBiddingFrame = pd.DataFrame(kxBiddingData, columns=province, index=kxSupplier)

kxData = {}
kxFrame = pd.DataFrame(kxData, columns=province, index=kxSupplier)

kxTotalPrice = {}
kxTotalPriceFrame = pd.DataFrame(kxTotalPrice, columns=province, index=kxSupplier)

kxOrderQuantity = {}
kxOrderQuantityFrame = pd.DataFrame(kxOrderQuantity, columns=province, index=kxSupplier)

# 空调（简称kt）
ktSqlStr = "select LEFT(o.province_name,2) '省分公司', p.real_nums '采购数量', " \
           "round(round(p.real_nums * p.company_price,5),2) AS '价税合计', " \
           "CASE " \
           "WHEN p.provider_id = 49216 THEN '美的制冷' " \
           "WHEN p.provider_id = 49280 THEN '广东科龙' " \
           "WHEN p.provider_id = 49289 THEN '南京佳力图' " \
           "WHEN p.provider_id = 49221 THEN '美的暖通' " \
           "WHEN p.provider_id = 49218 THEN '广东海悟' " \
           "WHEN p.provider_id = 49295 THEN '北京斯泰科' " \
           "WHEN p.provider_id = 49212 THEN '艾特网能' " \
           "WHEN p.provider_id = 49220 THEN '艾默生网络' " \
           "WHEN p.provider_id = 49258 THEN 'TCL空调' " \
           "WHEN p.provider_id = 49279 THEN '海信空调' " \
           "END '供应商' " \
           "from eshop_order_product p " \
           "LEFT JOIN eshop_order o ON p.order_id = o.id " \
           "LEFT JOIN eshop_provideraddress epa ON epa.providerId = p.provider_id " \
           "LEFT JOIN eshop_provider_contact c ON c.provider_id = p.provider_id " \
           "LEFT JOIN eshop_goods g ON g.item_all = p.ITEM_NUMBER " \
           "LEFT JOIN eshop_materials_catergorytree mac ON mac.id =p.goodstype_id " \
           "WHERE epa.shop_id = o.shop_id " \
           "AND c.shop_id = o.shop_id " \
           "AND g.shop_id = o.shop_id " \
           "AND o.shop_id = ' 596 ' " \
           "AND p.CONTACT_NUMBER = c.contact_number " \
           "AND p.CONTACT_NUMBER IN ('CU12-1001-2016-000977', 'CU12-1001-2016-000981', 'CU12-1001-2016-000982', " \
           "'CU12-1001-2016-000973', 'CU12-1001-2016-000978', 'CU12-1001-2016-000975', " \
           "'CU12-1001-2016-000980', 'CU12-1001-2016-000969', 'CU12-1001-2016-000972', " \
           "'CU12-1001-2016-000976', 'CU12-1001-2016-000983', 'CU12-1001-2016-000970', " \
           "'CU12-1001-2016-000974', 'CU12-1001-2016-000979', 'CU12-1001-2016-000971') " \
           "AND mac.name in ('节能减排类设备', '通信机房节能双循环空调') " \
           "AND p.unit = '台'" \
           "AND o.`status` in ('2','5') " \
           "AND o.create_time BETWEEN '2016-01-01' And '%s' " \
           "GROUP BY p.id " % end_data
ktInfo = "空调（单位：台）"
ktSupplier = ['广东海悟', '广东科龙', '海信空调', '美的制冷', 'TCL空调',
              '艾特网能', '北京斯泰科', '美的暖通', '南京佳力图', '艾默生网络']

ktTotalBiddingData = [4113, 3973, 3932, 3531, 3496, 759, 734, 724, 564, 203]

ktBiddingData = {'北京': [195, 199, 199, 170, 176, 19, 19, 20, 0, 19],
                 '天津': [175, 74, 73, 65, 64, 119, 117, 112, 104, 19],
                 '河北': [445, 418, 414, 369, 366, 90, 88, 83, 95, 0],
                 '山西': [19, 22, 21, 19, 19, 7, 6, 6, 7, 0],
                 '内蒙': [338, 359, 354, 323, 314, 31, 31, 30, 26, 0],
                 '辽宁': [117, 84, 126, 112, 112, 7, 7, 0, 7, 0],
                 '吉林': [26, 10, 10, 0, 0, 28, 28, 26, 30, 0],
                 '黑龙': [70, 30, 30, 27, 27, 45, 44, 43, 48, 0],
                 '上海': [227, 0, 6, 228, 258, 6, 234, 255, 0, 6],
                 '江苏': [108, 9, 19, 126, 123, 20, 112, 122, 5, 47],
                 '浙江': [152, 0, 0, 149, 172, 0, 156, 171, 0, 0],
                 '安徽': [66, 0, 0, 65, 75, 0, 69, 75, 1, 1],
                 '福建': [138, 0, 0, 135, 156, 0, 142, 155, 0, 0],
                 '江西': [27, 0, 0, 26, 30, 0, 27, 30, 4, 4],
                 '山东': [265, 32, 40, 299, 302, 40, 274, 299, 21, 21],
                 '河南': [163, 0, 39, 190, 177, 38, 156, 174, 42, 41],
                 '湖北': [66, 0, 10, 74, 71, 12, 63, 70, 11, 11],
                 '湖南': [226, 0, 36, 250, 246, 38, 218, 242, 39, 37],
                 '广东': [119, 79, 81, 194, 129, 85, 114, 127, 0, 80],
                 '广西': [128, 0, 0, 127, 139, 7, 124, 137, 0, 0],
                 '海南': [41, 0, 0, 38, 44, 0, 39, 44, 0, 0],
                 '重庆': [99, 12, 16, 112, 112, 17, 102, 111, 0, 13],
                 '四川': [44, 20, 30, 72, 47, 31, 42, 46, 11, 31],
                 '贵州': [125, 0, 9, 126, 135, 9, 120, 133, 10, 9],
                 '云南': [0, 0, 8, 8, 11, 8, 0, 10, 8, 8],
                 '西藏': [19, 0, 7, 24, 21, 7, 19, 21, 7, 7],
                 '陕西': [102, 13, 19, 113, 117, 14, 106, 115, 6, 19],
                 '甘肃': [75, 0, 36, 109, 53, 35, 75, 84, 39, 37],
                 '青海': [0, 0, 0, 0, 13, 0, 12, 13, 4, 4],
                 '宁夏': [0, 0, 6, 7, 87, 7, 12, 14, 4, 4],
                 '新疆': [228, 0, 32, 256, 259, 30, 230, 257, 35, 33]}
ktBiddingFrame = pd.DataFrame(ktBiddingData, columns=province, index=ktSupplier)

ktData = {}
ktFrame = pd.DataFrame(ktData, columns=province, index=ktSupplier)

ktTotalPrice = {}
ktTotalPriceFrame = pd.DataFrame(ktTotalPrice, columns=province, index=ktSupplier)

ktOrderQuantity = {}
ktOrderQuantityFrame = pd.DataFrame(ktOrderQuantity, columns=province, index=ktSupplier)

# 电源（简称dy）
# It looks like python is interpreting the % as a printf-like format character. Try using %%?
dySqlStr = "select LEFT(o.province_name,2) '省分公司', p.real_nums '采购数量', " \
           "round(round(p.real_nums * p.company_price,5),2) AS '价税合计', " \
           "CASE " \
           "WHEN p.provider_id = 49220 THEN '艾默生网络' " \
           "WHEN p.provider_id = 49291 THEN '珠江电信' " \
           "WHEN p.provider_id = 49259 THEN '南京华脉' " \
           "WHEN p.provider_id = 49213 THEN '中达电通' " \
           "WHEN p.provider_id = 49223 THEN '东莞铭普' " \
           "WHEN p.provider_id = 49229 THEN '北京动力源' " \
           "WHEN p.provider_id = 28295 THEN '华为公司' " \
           "WHEN p.provider_id = 49211 THEN '中兴公司' " \
           "END '供应商' " \
           "from eshop_order_product p " \
           "LEFT JOIN eshop_order o ON p.order_id = o.id " \
           "LEFT JOIN eshop_provideraddress epa ON epa.providerId = p.provider_id " \
           "LEFT JOIN eshop_provider_contact c ON c.provider_id = p.provider_id " \
           "LEFT JOIN eshop_goods g ON g.item_all = p.ITEM_NUMBER " \
           "LEFT JOIN eshop_materials_catergorytree mac ON mac.id =p.goodstype_id " \
           "WHERE epa.shop_id = o.shop_id " \
           "AND c.shop_id = o.shop_id " \
           "AND g.shop_id = o.shop_id " \
           "AND o.shop_id = ' 596 ' " \
           "AND p.CONTACT_NUMBER = c.contact_number " \
           "AND p.CONTACT_NUMBER IN ('CU12-1001-2016-000988','CU12-1001-2016-000995'," \
           "'CU12-1001-2016-000990','CU12-1001-2016-000987','CU12-1001-2016-000992'," \
           "'CU12-1001-2016-000989','CU12-1001-2016-000994','CU12-1001-2016-000998'," \
           "'CU12-1001-2016-000993','CU12-1001-2016-000991','CU12-1001-2016-000996','CU12-1001-2016-000997') " \
           "AND mac.name in ('组合开关电源', '室外一体化开关电源', '其他变流设备') " \
           "AND p.unit = '套'" \
           "AND p.Goods_name NOT LIKE '%%配件%%' " \
           "AND o.`status` in ('2','5') " \
           "AND o.create_time BETWEEN '2016-01-01' And '%s' " \
           "GROUP BY p.id " % end_data
dyInfo = "电源（单位：套）"
dySupplier = ['艾默生网络', '珠江电信', '南京华脉', '中达电通', '东莞铭普', '北京动力源', '华为公司', '中兴公司']

dyTotalBiddingData = [4066, 3957, 3745, 2998, 2292, 1641, 1244, 309]

dyBiddingData = {'北京': [12, 9, 0, 2, 0, 2, 2, 3],
                 '天津': [45, 33, 31, 44, 0, 43, 13, 13],
                 '河北': [617, 590, 563, 603, 0, 603, 50, 52],
                 '山西': [65, 63, 60, 63, 0, 62, 4, 5],
                 '内蒙': [48, 41, 40, 47, 0, 46, 8, 8],
                 '辽宁': [272, 246, 233, 267, 0, 263, 33, 36],
                 '吉林': [43, 41, 39, 43, 0, 42, 4, 4],
                 '黑龙': [195, 189, 179, 191, 0, 189, 14, 15],
                 '上海': [150, 154, 5, 149, 157, 146, 4, 5],
                 '江苏': [73, 75, 12, 62, 65, 61, 12, 13],
                 '浙江': [209, 214, 10, 204, 214, 200, 9, 10],
                 '安徽': [218, 222, 4, 219, 230, 215, 4, 4],
                 '福建': [174, 177, 2, 176, 185, 172, 0, 2],
                 '江西': [87, 89, 2, 87, 92, 85, 0, 2],
                 '山东': [254, 260, 20, 240, 252, 235, 19, 20],
                 '河南': [11, 151, 11, 138, 147, 135, 148, 11],
                 '湖北': [7, 54, 7, 47, 50, 46, 52, 7],
                 '湖南': [9, 177, 10, 167, 178, 163, 174, 10],
                 '广东': [11, 267, 11, 255, 271, 248, 262, 11],
                 '广西': [2, 151, 2, 149, 158, 144, 149, 2],
                 '海南': [7, 20, 7, 12, 13, 12, 18, 7],
                 '重庆': [35, 35, 2, 33, 35, 33, 0, 2],
                 '四川': [11, 39, 11, 28, 30, 27, 37, 11],
                 '贵州': [5, 132, 5, 127, 135, 125, 130, 5],
                 '云南': [14, 70, 14, 55, 58, 54, 68, 15],
                 '西藏': [3, 3, 2, 7, 8, 0, 2, 3],
                 '陕西': [195, 198, 4, 195, 205, 191, 4, 4],
                 '甘肃': [59, 61, 59, 44, 47, 0, 14, 16],
                 '青海': [12, 13, 0, 12, 12, 0, 0, 0],
                 '宁夏': [2, 12, 2, 10, 10, 0, 0, 2],
                 '新疆': [190, 195, 189, 184, 193, 0, 10, 11]}
dyBiddingFrame = pd.DataFrame(dyBiddingData, columns=province, index=dySupplier)

dyData = {}
dyFrame = pd.DataFrame(dyData, columns=province, index=dySupplier)

dyTotalPrice = {}
dyTotalPriceFrame = pd.DataFrame(dyTotalPrice, columns=province, index=dySupplier)

dyOrderQuantity = {}
dyOrderQuantityFrame = pd.DataFrame(dyOrderQuantity, columns=province, index=dySupplier)

# 微基站（简称wjz）
wjzSqlStr = "select LEFT(o.province_name,2) '省分公司', p.real_nums '采购数量', " \
            "round(round(p.real_nums * p.company_price,5),2) AS '价税合计', " \
            "CASE " \
            "WHEN p.provider_id = 37708 THEN '爱立信公司' " \
            "WHEN p.provider_id IN ('37771', '37715') THEN '上海贝尔' " \
            "WHEN p.provider_id IN ('37623','37625') THEN '华为公司' " \
            "WHEN p.provider_id IN ('37593','37626') THEN '中兴公司' " \
            "END '供应商' " \
            "from eshop_order_product p " \
            "LEFT JOIN eshop_order o ON p.order_id = o.id " \
            "LEFT JOIN eshop_provideraddress epa ON epa.providerId = p.provider_id " \
            "LEFT JOIN eshop_provider_contact c ON c.provider_id = p.provider_id " \
            "LEFT JOIN eshop_goods g ON g.item_all = p.ITEM_NUMBER " \
            "LEFT JOIN eshop_materials_catergorytree mac ON mac.id =p.goodstype_id " \
            "WHERE epa.shop_id = o.shop_id " \
            "AND c.shop_id = o.shop_id " \
            "AND g.shop_id = o.shop_id " \
            "AND o.shop_id = ' 685 ' " \
            "AND p.CONTACT_NUMBER = c.contact_number " \
            "AND p.CONTACT_NUMBER IN ('CU12-1001-2016-000735','CU12-1001-2016-000729','CU12-1001-2016-000730'," \
            "'CU12-1001-2016-000732','CU12-1001-2016-000736','CU12-1001-2016-000734'," \
            "'CU12-1001-2016-000731','CU12-1001-2016-000733') " \
            "AND mac.name in ('pPRRU') " \
            "AND o.`status` in ('2','5') " \
            "AND o.create_time BETWEEN '2016-01-01' And '%s' " \
            "GROUP BY p.id " % end_data
wjzInfo = "微基站（单位：个）"
wjzSupplier = ['华为公司', '中兴公司', '爱立信公司', '上海贝尔']

wjzTotalBiddingData = [8376, 6369, 4001, 570]

wjzBiddingData = {'北京': [1128, 0, 778, 0],
                  '天津': [48, 150, 0, 0],
                  '河北': [48, 30, 0, 100],
                  '山西': [2208, 1261, 450, 0],
                  '内蒙': [432, 0, 0, 117],
                  '辽宁': [48, 8, 0, 233],
                  '吉林': [48, 195, 32, 0],
                  '黑龙': [48, 20, 0, 13],
                  '上海': [24, 0, 0, 50],
                  '江苏': [48, 319, 32, 0],
                  '浙江': [912, 830, 0, 0],
                  '安徽': [48, 312, 0, 13],
                  '福建': [48, 406, 0, 0],
                  '江西': [24, 0, 0, 0],
                  '山东': [336, 819, 250, 0],
                  '河南': [1320, 465, 0, 0],
                  '湖北': [48, 46, 490, 0],
                  '湖南': [48, 400, 0, 10],
                  '广东': [48, 35, 144, 0],
                  '广西': [48, 0, 0, 7],
                  '海南': [0, 14, 8, 0],
                  '重庆': [432, 132, 0, 0],
                  '四川': [144, 236, 1812, 10],
                  '贵州': [312, 0, 0, 0],
                  '云南': [48, 0, 5, 0],
                  '西藏': [0, 0, 0, 0],
                  '陕西': [48, 60, 0, 17],
                  '甘肃': [48, 119, 0, 0],
                  '青海': [24, 5, 0, 0],
                  '宁夏': [48, 102, 0, 0],
                  '新疆': [312, 405, 0, 0]}
wjzBiddingFrame = pd.DataFrame(wjzBiddingData, columns=province, index=wjzSupplier)

wjzData = {}
wjzFrame = pd.DataFrame(wjzData, columns=province, index=wjzSupplier)

wjzTotalPrice = {}
wjzTotalPriceFrame = pd.DataFrame(wjzTotalPrice, columns=province, index=wjzSupplier)

wjzOrderQuantity = {}
wjzOrderQuantityFrame = pd.DataFrame(wjzOrderQuantity, columns=province, index=wjzSupplier)

# 普通光缆（简称ptgl）
ptglSqlStr = ""
ptglInfo = "微基站（单位：个）"
ptglSupplier = ['华为公司', '中兴公司', '爱立信公司', '上海贝尔']

ptglTotalBiddingData = [8376, 6369, 4001, 570]


#  格式
#  表头格式
header_style = xlwt.easyxf('font: name 微软雅黑, height 220, bold on;')
#  表行列名格式
tablestyle = 'font: name 微软雅黑, height 180, bold on; '  # 粗体字
tablestyle += 'align: horz centre, vert center, wrap on; '  # 居中,自动换行
tablestyle += 'borders: left THIN, right THIN, top THIN, bottom THIN; '  # 边框
table_style = xlwt.easyxf(tablestyle)
#  正文格式
textstyle = 'font: name 微软雅黑, height 180;'  # 粗体字
textstyle += 'align: horz centre, vert center, wrap on; '  # 居中
textstyle += 'borders: left THIN, right THIN, top THIN, bottom THIN; '  # 边框
text_style = xlwt.easyxf(textstyle)
# 百分比格式
percent_font = xlwt.Font()
percent_font.name = '微软雅黑'
percent_font.height = 180
percent_borders = xlwt.Borders()
percent_borders.left = xlwt.Borders.THIN
percent_borders.right = xlwt.Borders.THIN
percent_borders.top = xlwt.Borders.THIN
percent_borders.bottom = xlwt.Borders.THIN
percent_style = xlwt.XFStyle()
percent_style.num_format_str = '0.00%'
percent_style.font = percent_font
percent_style.borders = percent_borders

import pandas as pd
import xlwt
import pymysql.cursors
import datetime

# 截至日
end_data = datetime.date(2017, 7, 31)

#  路径前加r（原因：文件名中的 \U 开始的字符被编译器认为是八进制）
#  保存输出数据的文档地址  Administrator
# resultFile_path = r"C:\Users\Administrator\Desktop\数据结果.xls"
resultFile_path = r"C:\Users\Zhang Yu\Desktop\数据结果.xls"

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
dldlSqlStr = "select LEFT(o.province_name,2) '省分公司', p.real_nums '采购数量', " \
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
             "GROUP BY p.id " %end_data


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

# 馈线（简称kx）
kxInfo = "馈线（单位：米）"
kxSupplier = ['江苏俊知', '江苏亨鑫', '珠海汉胜', '长飞光纤', '通鼎互联',
              '中天射频', '成都大唐', '富通集团', '湖北凯乐']

kxTotalBiddingData = [41599100, 41117902, 39462841, 21376628, 21081553, 10225762, 10036104, 7042699, 7016808]

kxBiddingData = {'北京': [17151, 16500, 16443, 0, 0, 15881, 15425, 0, 0],
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
kxBiddingFrame = pd.DataFrame(kxBiddingData, columns=province, index=kxSupplier)

kxData = {}
kxFrame = pd.DataFrame(kxData, columns=province, index=kxSupplier)

#  格式
#  表头格式
header_style = xlwt.easyxf('font: name 微软雅黑, height 220, bold on;')
#  表行列名格式
tablestyle = 'font: name 微软雅黑, height 180, bold on; '  #  粗体字
tablestyle += 'align: horz centre, vert center, wrap on; '  #  居中,自动换行
tablestyle += 'borders: left THIN, right THIN, top THIN, bottom THIN; '  #  边框
table_style = xlwt.easyxf(tablestyle)
#  正文格式
textstyle = 'font: name 微软雅黑, height 180;'  #  粗体字
textstyle += 'align: horz centre, vert center, wrap on; '  #  居中
textstyle += 'borders: left THIN, right THIN, top THIN, bottom THIN; '  #  边框
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

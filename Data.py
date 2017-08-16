import pandas as pd
# 所有省分
province = ['北京', '天津', '河北', '山西', '内蒙', '辽宁', '吉林', '黑龙', '上海', '江苏', '浙江',
            '安徽', '福建', '江西', '山东', '河南', '湖北', '湖南', '广东', '广西', '海南', '重庆',
            '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海', '宁夏', '新疆']

# 电力电缆（简称dldl）
dldlInfo = "电力电缆（单位：米）备注：江苏中利集团股份有限即中利科技"
dldlSupplier = ['江苏亨通', '中天科技', '江苏俊知', '成都大唐', '富通集团', '通鼎互联', '中利科技', '鲁能泰山', '西部电缆']

dldlTotalBiddingData = [17151, 16500, 16443, 0, 0, 15881, 15425, 0, 0]

dldlBiddingData = {'北京': [17151, 16500, 16443, 0, 0, 15881, 15425, 0, 0]}
dldlBiddingFrame = pd.DataFrame(dldlBiddingData, columns=province, index=dldlSupplier)

dldlData = {'北京': [0, 0, 0, 0, 0, 0, 0, 0, 0]}
dldlFrame = pd.DataFrame(dldlData, columns=province, index=dldlSupplier)

# dldlFrame.iloc[dldlSupplier.index('江苏亨通'), province.index('天津')] = 100
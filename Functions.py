import math
import Data

def totalCount(list):
    result = 0
    for n in range(0, len(list)):
        result += list[n]
    return result

def percentTran(int):
    return format(int, '.0%')
    # return str(round((int*100), 2))+'%'

def intTran(dataframeloc):
    return 0 if math.isnan(dataframeloc) else int(dataframeloc)

def totalCountByProvince(m, dataframe):
    return (0 if math.isnan(dataframe[Data.province[m]].T.sum())
            else int(dataframe[Data.province[m]].T.sum()))

def mergeStr(n):
    return 'I'+str(n)+'+L'+str(n)+'+O'+str(n)+'+R'+str(n)+'+U'+str(n)+'+X'+str(n)+'+AA'+str(n)+'+AD'+str(n)+'+AG'+str(n)+'+AJ'+str(n)+'+AM'\
           +str(n)+'+AP'+str(n)+'+AS'+str(n)+'+AV'+str(n)+'+AY'+str(n)+'+BB'+str(n)+'+BE'+str(n)+'+BH'+str(n)+'+BK'+str(n)+'+BN'+str(n)+'+BQ'\
           +str(n)+'+BT'+str(n)+'+BW'+str(n)+'+BZ'+str(n)+'+CC'+str(n)+'+CF'+str(n)+'+CI'+str(n)+'+CL'+str(n)+'+CO'+str(n)+'+CR'+str(n)+'+CU'+str(n)

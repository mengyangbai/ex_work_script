import tushare as ts
 
# 获取实时行情数据
hq = ts.get_today_all()
# 节选出股票代码code、名称name、涨跌幅changepercent、股价trade
hq = hq[['code','name','changepercent','trade']]
 
# 筛选出当前股价高于0元低于3元的股票信息
mins = hq.trade>0.00
maxs = hq.trade<=2.99
allselect = mins & maxs
data = hq[allselect].sort('trade')
import sys
sys.path.append('.')
import tdxtrader
#创建user
user = pytdxtrader.TdxTrader('1234','hbzq')
#当日委托
print(user.update_order()) 
#当日成交
print(user.update_trade())
#发单, 参数(合约代码，方向，价格，数量)
print(user.send_order('600000',0,10,100))
#撤单,参数指定委托号
#user.cancel_order('2881')
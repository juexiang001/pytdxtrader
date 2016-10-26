# pytdxtrader

manipulate tdx by python.

#开发环境: Windows 7 / Python 3.5

# 讨论
  QQ群277699808
  

#通达信版本:v6
  现支持 华宝证券，西南证券，安信证券 通达信
  添加新券商支持的方法，参考 tdxtrader.py config的配置方法,主要是添加当日委托列表，当日成交列表两个listview的列信息（因为不同券商配置不一样)

#requirements

pip install -r requirements.txt

#用法

#引入:

先启动通达信，并登陆交易账户，再运行python

import pytdxtrader


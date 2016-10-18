import win32api
import win32con
import win32gui
import win32process
print('test3')
import time as tm
import pywinauto

class TdxTrader:
    od_long = 0
    od_short = 1

    title_re_ = ''
    account_id = ''
    hTdxMain = 0
    hTreeLeft = 0
    hRight = 0
    hDlgSend = 0
    hOrderList = 0
    hOrderRefresh = 0
    lOrderColumn = []
    hTradeList = 0
    hTradeRefresh = 0
    lTradeColumn =[]
    
    hBuySymbol = 0
    hBuyPrice = 0
    hBuyVol = 0
    hBuyCmd = 0
    hBuyDirection = 0

    hSellSymbol = 0
    hSellPrice = 0
    hSellVol = 0
    hSellCmd = 0
    hSellDirection = 0
    
    threadId = 0
    processId = 0
    nOrderList = 0
    nTradeList = 0
    
    account_info_ = {'name':''}
    
    order_status_str_ = {}
    
    config = {}
    config['hbzq'] = {'lOrderColumn':[{'text':'委托时间','key':'time'},\
    {'text':'股东代码','key':'account_id'}\
    ,{'text':'证券代码','key':'symbol'}\
    ,{'text':'证券名称','key':'name'}\
    ,{'text':'委托类别','key':'direction'}\
    ,{'text':'委托价格','key':'price'}\
    ,{'text':'委托数量','key':'volume'}\
    ,{'text':'成交价格','key':'traded_price'}\
    ,{'text':'成交数量','key':'volume_traded'}\
    ,{'text':'撤单数量','key':'volume_canceled'}\
    ,{'text':'委托编号','key':'exchange_order_id'}\
    ,{'text':'报价方式','key':'price_type'}\
    ,{'text':'状态说明','key':'status'}\
    ,{'text':'冻结资金','key':'frozen_margin'}]\
    ,\
    'lTradeColumn':\
    [{'text':'成交时间','key':'time'}\
    ,{'text':'股东代码','key':'account_id'}\
    ,{'text':'证券代码','key':'symbol'}\
    ,{'text':'证券名称','key':'name'}\
    ,{'text':'委托类别','key':'direction'}\
    ,{'text':'成交价格','key':'price'}\
    ,{'text':'成交数量','key':'volume'}\
    ,{'text':'发生金额','key':'amount'}\
    ,{'text':'成交编号','key':'trade_id'}\
    ,{'text':'委托编号','key':'exchange_order_id'}]\
    ,\
    'title_re':'华宝'\
    ,'order_status_str':{'done':'已成交','canceled':'已撤单已撤','pending':'已申报'}\
    }
    
    config['xnzq']={\
    'lOrderColumn':[{'text':'委托时间','key':'time'},{'text':'证券代码','key':'symbol'},\
    {'text':'证券名称','key':'name'},{'text':'买卖标志','key':'direction'},{'text':'委托类别','key':'order_type'}\
    ,{'text':'状态说明','key':'status'},{'text':'委托价格','key':'price'},{'text':'委托数量','key':'volume'}\
    ,{'text':'委托编号','key':'exchange_order_id'},{'text':'成交价格','key':'traded_price'},{'text':'成交数量','key':'volume_traded'}\
    ,{'text':'报价方式','key':'price_type'},{'text':'股东代码','key':'account_id'},{'text':'废单原因','key':'error_info'}]\
    ,\
    'lTradeColumn':[{'text':'成交时间','key':'time'},{'text':'证券代码','key':'symbol'}\
    ,{'text':'证券名称','key':'name'},{'text':'买卖标志','key':'direction'},{'text':'成交价格','key':'price'}\
    ,{'text':'成交数量','key':'volume'},{'text':'成交金额','key':'amount'},{'text':'成交编号','key':'trade_id'}\
    ,{'text':'委托编号','key':'exchange_order_id'},{'text':'股东代码','key':'account_id'},{'text':'成交类型','key':'trade_type'}]\
    ,\
    'title_re':'西南'\
    ,'order_status_str':{'done':'已成交','canceled':'已撤单已撤','pending':'已申报'}\
    }
    
    config['axzq']={\
    'lOrderColumn':[{'text':'委托时间','key':'time'},{'text':'委托编号','key':'exchange_order_id'}\
    ,{'text':'证券代码','key':'symbol'},{'text':'证券名称','key':'name'},{'text':'买卖标志','key':'direction'}\
    ,{'text':'委托类型','key':'order_type'},{'text':'委托价格','key':'price'},{'text':'委托数量','key':'volume'},{'text':'委托金额','key':'amount'}\
    ,{'text':'成交价格','key':'traded_price'},{'text':'成交数量','key':'volume_traded'},{'text':'成交金额','key':'traded_amount'}\
    ,{'text':'已撤数量','key':'volume_canceled'},{'text':'撤单标志','key':'status'},{'text':'股东代码','key':'account_id'}\
    ,{'text':'市场','key':'exchange'},{'text':'客户代码','key':'investorid'},{'text':'资金帐号','key':'account_id'}\
    ,{'text':'股东姓名','key':'account_name'},{'text':'合法标志','key':'legal_flag'},{'text':'摘要','key':'comment'}\
    ],\
    'lTradeColumn':[{'text':'委托时间','key':'time'},{'text':'委托编号','key':'exchange_order_id'},{'text':'证券代码','key':'symbol'}\
    ,{'text':'证券名称','key':'name'},{'text':'买卖标志','key':'direction'},{'text':'委托类型','key':'order_type'}\
    ,{'text':'委托价格','key':'order_price'},{'text':'委托数量','key':'order_volume'},{'text':'成交价格','key':'price'}\
    ,{'text':'成交数量','key':'volume'},{'text':'成交金额','key':'amount'},{'text':'成交时间','key':'time'}\
    ,{'text':'股东代码','key':'holder_id'},{'text':'资金帐号','key':'account_id'},{'text':'客户代码','key':'investorid'}\
    ,{'text':'股东姓名','key':'account_name'},{'text':'市场','key':'exchange'},{'text':'成交编号','key':'trade_id'}\
    ],\
    'title_re':'安信'\
    ,'order_status_str':{'done':'已成交','canceled':'已撤单已撤','pending':'已申报'}\
    }

    def __init__(self, id, broker):
        if broker == 'axtdx':
            broker = 'axzq'
        elif broker == 'hbtdx':
            broker == 'hbzq'
        elif broker == 'xntdx':
            broker = 'xnzq'
        elif broker == '':
            broker = 'hbzq'
        self.account_id = id
        self.lOrderColumn = self.config[broker]['lOrderColumn']
        self.lTradeColumn = self.config[broker]['lTradeColumn']
        self.title_re_ = self.config[broker]['title_re']
        self.order_status_str_ = self.config[broker]['order_status_str']
        print('start get handle')
        self.account_info_['account_id'] = id
        self.gethandle()
        
    def check_is_login(self):
        if self.nOrderList > 1 or self.nTradeList > 1 or self.nOrderList == 0 or self.nTradeList == 0:
            return False
        if self.hOrderList == 0 or self.hTradeList == 0 or self.hOrderRefresh == 0 or self.hTradeRefresh == 0\
        or self.hBuySymbol == 0 or self.hBuyPrice == 0 or self.hBuyVol == 0 or self.hBuyCmd == 0 or self.hSellSymbol == 0\
        or self.hSellPrice == 0 or self.hSellVol == 0 or self.hSellCmd == 0 or self.hBuyDirection == 0 or self.hSellDirection == 0:
            return False
        return True
    def check_foreground(self):
        try:
            win32gui.ShowWindow(self.hTdxMain,win32con.SW_NORMAL)
            win32gui.SetActiveWindow(self.hTdxMain)
            win32gui.SetForegroundWindow(self.hTdxMain)
        except:
            print('check_foreground except')

        for i in range(1,10):
            if pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).IsVisible() == False:
                print('check_foreground')
                tm.sleep(0.1)
            else:
                break
                
    def map_order_status(self, od):
        if self.order_status_str_['done'].count(od['status']) > 0:
            od['status'] = 0
        elif self.order_status_str_['canceled'].count(od['status']) > 0:
            od['status'] = 4
        elif self.order_status_str_['pending'].count(od['status']) > 0:
            od['status'] = 3
        else:
            od['status'] = 5

    def gethandle(self):
        #获取所有控件句柄
        hwnd1 = pywinauto.findwindows.find_window(class_name='TdxW_MainFrame_Class',title_re=self.title_re_)
        if hwnd1 == 0 :
            return False
        self.hTdxMain = hwnd1
        win32gui.ShowWindow(self.hTdxMain,win32con.SW_NORMAL)
        win32gui.SetActiveWindow(self.hTdxMain)
        win32gui.SetForegroundWindow(self.hTdxMain)
        self.threadId, self.processId = win32process.GetWindowThreadProcessId(hwnd1)
        hX = win32gui.GetDlgItem(hwnd1, 0xE81E)
        hX = win32gui.GetDlgItem(hX, 0xF5)
        hDlg = win32gui.FindWindowEx(hX,0,'AfxWnd42','')
        hDlg = win32gui.FindWindowEx(hX,hDlg,'AfxWnd42','')
        hDlg = win32gui.FindWindowEx(hX,hDlg,'AfxWnd42','')
        hDlg = win32gui.FindWindowEx(hX,hDlg,'AfxWnd42','')
        hDlg = win32gui.FindWindowEx(hDlg,0,"#32770",'通达信网上交易V6')
        if pywinauto.controls.HwndWrapper.HwndWrapper(hDlg).IsVisible() == False:
            win32gui.PostMessage(hwnd1, win32con.WM_KEYDOWN, win32con.VK_F12)
            win32gui.PostMessage(hwnd1, win32con.WM_KEYUP, win32con.VK_F12)
            tm.sleep(0.2)
            print('show F12')
        hDlg = win32gui.GetDlgItem(hDlg, 0)
        hDlg = win32gui.FindWindowEx(hDlg,0,"AfxMDIFrame42",'')
        hLeft = win32gui.GetDlgItem(hDlg, 0xE900)
        hTreeLeft = win32gui.GetDlgItem(hLeft, 0xDD)
        self.hTreeLeft = win32gui.GetDlgItem(hTreeLeft, 0xE900)
        self.hRight = win32gui.GetDlgItem(hDlg, 0xE901)
        hRightHeader = win32gui.GetDlgItem(self.hRight, 0xE81B) 
        hRightHeader = win32gui.GetDlgItem(hRightHeader, 0) ;print('hRightHeader:%x'%hRightHeader)
        hAccountInfo = win32gui.FindWindowEx(hRightHeader,0,'ComboBox','')
        print(pywinauto.controls.HwndWrapper.HwndWrapper(hAccountInfo).WindowText())
        self.close_all_dialog()
        pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).SetFocus()
        path = "\查询\当日委托";
        pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).EnsureVisible(path)
        rect = pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).GetItem(path).Rectangle()
        pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).ClickInput(coords=(rect.left+5,rect.top+5));tm.sleep(1);
        
        pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).SetFocus()
        path = "\查询\当日成交"
        pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).EnsureVisible(path)
        rect = pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).GetItem(path).Rectangle()
        pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).ClickInput(coords=(rect.left+5,rect.top+5));tm.sleep(1);
        
        pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).SetFocus()
        path = "\对买对卖"
        pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).EnsureVisible(path)
        rect = pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).GetItem(path).Rectangle()
        pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).ClickInput(coords=(rect.left+5,rect.top+5));tm.sleep(1);

        hDlg = 0
        self.hOrderList = 0
        self.hTradeList = 0
        self.hOrderRefresh = 0
        self.hTradeRefresh = 0
        self.hBuySymbol = 0
        self.hBuyPrice = 0
        self.hBuyVol = 0
        self.hBuyCmd = 0
        self.hSellSymbol = 0
        self.hSellPrice = 0
        self.hSellVol = 0
        self.hSellCmd = 0
        
        self.nOrderList = 0
        self.nTradeList = 0
        for i in range(1,30):
            hDlg = win32gui.FindWindowEx(self.hRight,hDlg,'#32770','')
            if hDlg == 0:
                break;
            try:
                hB1 = win32gui.GetDlgItem(hDlg, 0x4E3)
                if hB1 > 0:
                    if pywinauto.controls.HwndWrapper.HwndWrapper(hB1).WindowText() == '买卖关联同一支股票':
                        hChk = win32gui.GetDlgItem(hDlg, 0x485)
                        pywinauto.controls.HwndWrapper.HwndWrapper(hChk).UnCheck()
                        self.hBuySymbol = win32gui.GetDlgItem(hDlg, 0x2EE5)
                        self.hBuyPrice = win32gui.GetDlgItem(hDlg, 0x2EE6)
                        self.hBuyVol = win32gui.GetDlgItem(hDlg, 0x2EE7)
                        self.hBuyCmd = win32gui.GetDlgItem(hDlg, 0x7DA)
                        self.hBuyDirection = win32gui.GetDlgItem(hDlg, 0x2F00)
                        self.hSellSymbol = win32gui.GetDlgItem(hDlg, 0x7E9)
                        self.hSellPrice = win32gui.GetDlgItem(hDlg, 0x2F07)
                        self.hSellVol = win32gui.GetDlgItem(hDlg, 0xBD6)
                        self.hSellCmd = win32gui.GetDlgItem(hDlg, 0xBD8)
                        self.hSellDirection = win32gui.GetDlgItem(hDlg, 0xBE9)
                        print('get SendOrder handle')
            except:
                hDlg = hDlg
            finally:
                hDlg = hDlg
            try:
                hL1 = win32gui.GetDlgItem(hDlg, 0x61F)
                if hL1 > 0:
                    clist = pywinauto.controls.common_controls.ListViewWrapper(hL1).Columns()
                    isOrderList = 1
                    if len(clist) >= len(self.lOrderColumn):
                        for j in range(len(self.lOrderColumn)):
                            if clist[j]['text'] != self.lOrderColumn[j]['text']:
                                isOrderList = 0
                                print('tdx:%s, config:%s'%(clist[j]['text'],self.lOrderColumn[j]['text']))
                                break;
                    else:
                        isOrderList = 0
                    if isOrderList == 1:
                        self.nOrderList = self.nOrderList + 1
                        self.hOrderList = hL1
                        self.hOrderRefresh = win32gui.GetDlgItem(hDlg, 0x474)
                        print('get OrderListView')
                    isTradeList = 1
                    if len(clist) >= len(self.lTradeColumn):
                        for k in range(0, len(self.lTradeColumn)):
                            if clist[k]['text'] != self.lTradeColumn[k]['text']:
                                isTradeList = 0
                                print('tdx:%s, config:%s'%(clist[k]['text'],self.lTradeColumn[k]['text']))
                                break;
                    else:
                        isTradeList = 0
                    if isTradeList == 1:
                        self.nTradeList = self.nTradeList + 1
                        self.hTradeList = hL1
                        self.hTradeRefresh = win32gui.GetDlgItem(hDlg, 0x474)
                        print('get TradeListView')
            except:
                hDlg = hDlg
            finally:
                hDlg = hDlg
        print('orderListCount:%d , self.nTradeList:%d'%(self.nOrderList,self.nTradeList))
        return 0
        
    def update_order(self):
        print('update_order')
        self.close_all_dialog()
        if self.check_is_login() == False:
            self.gethandle()
            
        if self.hTreeLeft == 0 or self.hOrderList == 0:
            return []
        for i in range(1,10):
            self.check_foreground()
            if pywinauto.controls.win32_controls.ButtonWrapper(self.hOrderRefresh).IsVisible() == False:
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).SetFocus()
                path = "\查询\当日委托"
                pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).EnsureVisible(path)
                rect = pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).GetItem(path).Rectangle()
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).ClickInput(coords=(rect.left+5,rect.top+5))
                print('click order')
                tm.sleep(0.2)
            else:
                break
        t = pywinauto.controls.common_controls.ListViewWrapper(self.hOrderList).Texts()
        cc = len(self.lOrderColumn)
        ret = []
        if len(t) > 0 and (len(t)-1)%cc == 0:
            for i in range(1,len(t)):
                r = int((i-1)/cc)
                c = (i-1)%cc
                if len(ret) <= r :
                    ret.append({'comment':''})
                ret[r][self.lOrderColumn[c]['key']] = t[i];#print('r%d, key:%s, value:%s'%(r,self.lOrderColumn[c]['key'],t[i]))
        for i in range(0,len(ret)):
            self.map_order_status(ret[i])
        
        return ret
        
    def update_trade(self):
        print('update_trade')
        self.close_all_dialog()
        if self.check_is_login() == False:
            self.gethandle()

        if self.hTreeLeft == 0 or self.hTradeList == 0 :
            return []
        for i in range(1,10):
            self.check_foreground()
            if pywinauto.controls.win32_controls.ButtonWrapper(self.hTradeRefresh).IsVisible() == False:
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).SetFocus()
                path = "\查询\当日成交"
                pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).EnsureVisible(path)
                rect = pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).GetItem(path).Rectangle()
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).ClickInput(coords=(rect.left+5,rect.top+5))
                print('click trade')
                tm.sleep(0.2)
            else:
                break
        t = pywinauto.controls.common_controls.ListViewWrapper(self.hTradeList).Texts()
        cc = len(self.lTradeColumn)
        ret = []
        if len(t) > 0 and (len(t)-1)%cc == 0:
            for i in range(1,len(t)):
                r = int((i-1)/cc)
                c = (i-1)%cc
                if len(ret) <= r :
                    ret.append({})
                ret[r][self.lTradeColumn[c]['key']] = t[i]
        return ret
    def buy(self , symbol, direction , price , volume):
        if pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuyCmd).WindowText() != '买入确认':
            return {'exchange_order_id':'','error_info':'hBuyCmd方向不是买入'}
        input_result = 0
        for m in range(1,10):
            pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuySymbol).SetEditText(symbol)
            pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuyPrice).SetEditText(str(price))
            pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuyVol).SetEditText(str(volume))
            if pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuySymbol).WindowText() == symbol and \
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuyPrice).WindowText() == str(price) and\
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuyVol).WindowText() == str(volume) :                    
                for i in range(1,30):
                    if(pywinauto.controls.win32_controls.ButtonWrapper(self.hBuyCmd).IsEnabled()):
                        pywinauto.controls.win32_controls.ButtonWrapper(self.hBuyCmd).Click()
                        input_result = 1
                        break
                    else:
                        tm.sleep(0.1)
                break
            else:
                tm.sleep(0.3)
        if input_result == 0:
            return {'exchange_order_id':'','error_info':'输入参数未完成'}
        hResult = 0
        for i in range(1,15):
            try:
                hResult = pywinauto.findwindows.find_window(class_name='#32770',title_re='交易确认',process = self.processId)
                hResultText = win32gui.GetDlgItem(hResult, 0x1B65)
                hResultConfirm = win32gui.GetDlgItem(hResult, 0x1B67)
                t = pywinauto.controls.HwndWrapper.HwndWrapper(hResultText).WindowText().splitlines()
                if len(t) >= 4:
                    if t[0].count('买入') == 0 or t[1].count(symbol) == 0 or t[2].count(str(price)) == 0 or t[3].count(str(volume)) == 0 :
                        return {'exchange_order_id':'','error_info':'交易确认参数错误'}
                pywinauto.controls.HwndWrapper.HwndWrapper(hResultConfirm).Click()
                break
            except:
                tm.sleep(0.1)
        if hResult == 0:
            return {'exchange_order_id':'','error_info':'交易确认未正常'}
        hResult = 0
        for i in range(1,30):
            try:
                hResult = pywinauto.findwindows.find_window(class_name='#32770',title_re='提示',process = self.processId)
                hResultText = win32gui.GetDlgItem(hResult, 0x1B65)
                hResultConfirm = win32gui.GetDlgItem(hResult, 0x1B67)
                t = pywinauto.controls.HwndWrapper.HwndWrapper(hResultText).WindowText().splitlines()
                pywinauto.controls.HwndWrapper.HwndWrapper(hResultConfirm).Click()
                if len(t) >= 2 and t[0].count('原因') :
                    return {'exchange_order_id':'','error_info':t[1]}
                elif len(t) == 1 and t[0].count('合同号是') :
                    xr = t[0].partition('合同号是')
                    if len(xr) == 3:
                        return {'exchange_order_id':xr[2],'error_info':''}
                print(t)
                break
            except:
                tm.sleep(0.1)
        if hResult == 0:
            return {'exchange_order_id':'','error_info':'交易结果未正常'}
            
            
    def sell(self , symbol, direction , price , volume):
        if pywinauto.controls.HwndWrapper.HwndWrapper(self.hSellCmd).WindowText() != '卖出确认':
            return {'exchange_order_id':'','error_info':'hSellCmd方向不是卖出'}
        input_result = 0
        for m in range(1,10):
            pywinauto.controls.HwndWrapper.HwndWrapper(self.hSellSymbol).SetEditText(symbol)
            pywinauto.controls.HwndWrapper.HwndWrapper(self.hSellPrice).SetEditText(str(price))
            pywinauto.controls.HwndWrapper.HwndWrapper(self.hSellVol).SetEditText(str(volume))
            if pywinauto.controls.HwndWrapper.HwndWrapper(self.hSellSymbol).WindowText() == symbol and \
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hSellPrice).WindowText() == str(price) and\
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hSellVol).WindowText() == str(volume) :                    
                for i in range(1,30):
                    if(pywinauto.controls.win32_controls.ButtonWrapper(self.hSellCmd).IsEnabled()):
                        pywinauto.controls.win32_controls.ButtonWrapper(self.hSellCmd).Click()
                        input_result = 1
                        break
                    else:
                        tm.sleep(0.1)
                break
            else:
                tm.sleep(0.3)
        if input_result == 0:
            return {'exchange_order_id':'','error_info':'输入参数未完成'}
        hResult = 0
        for i in range(1,15):
            try:
                hResult = pywinauto.findwindows.find_window(class_name='#32770',title_re='交易确认',process = self.processId)
                hResultText = win32gui.GetDlgItem(hResult, 0x1B65)
                hResultConfirm = win32gui.GetDlgItem(hResult, 0x1B67)
                t = pywinauto.controls.HwndWrapper.HwndWrapper(hResultText).WindowText().splitlines()
                if len(t) >= 4:
                    if t[0].count('卖出') == 0 or t[1].count(symbol) == 0 or t[2].count(str(price)) == 0 or t[3].count(str(volume)) == 0 :
                        return {'exchange_order_id':'','error_info':'交易确认参数错误'}
                pywinauto.controls.HwndWrapper.HwndWrapper(hResultConfirm).Click()
                break
            except:
                tm.sleep(0.1)
        if hResult == 0:
            return {'exchange_order_id':'','error_info':'交易确认未正常'}
        hResult = 0
        for i in range(1,30):
            try:
                hResult = pywinauto.findwindows.find_window(class_name='#32770',title_re='提示',process = self.processId)
                hResultText = win32gui.GetDlgItem(hResult, 0x1B65)
                hResultConfirm = win32gui.GetDlgItem(hResult, 0x1B67)
                t = pywinauto.controls.HwndWrapper.HwndWrapper(hResultText).WindowText().splitlines()
                pywinauto.controls.HwndWrapper.HwndWrapper(hResultConfirm).Click()
                if len(t) >= 2 and t[0].count('原因') :
                    return {'exchange_order_id':'','error_info':t[1]}
                elif len(t) == 1 and t[0].count('合同号是') :
                    xr = t[0].partition('合同号是')
                    if len(xr) == 3:
                        return {'exchange_order_id':xr[2],'error_info':''}
                print(t)
                break
            except:
                tm.sleep(0.3)
        if hResult == 0:
            return {'exchange_order_id':'','error_info':'交易结果未正常'}
            
    def send_order(self , symbol, direction , price , volume):
        print('SendOrder:%s,%d,%s,%s'%(symbol,direction,str(price),str(volume)))
        self.close_all_dialog()
        if self.check_is_login() == False:
            self.gethandle()
        ret = {'exchange_order_id':'','error_info':''}
        if self.hTreeLeft == 0:
            return {'exchange_order_id':'','error_info':'handle=0'}
        for i in range(1,10):
            self.check_foreground()
            if pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuyCmd).IsVisible() == False:
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).SetFocus()
                path = "\对买对卖"
                pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).EnsureVisible(path)
                rect = pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).GetItem(path).Rectangle()
                pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).ClickInput(coords=(rect.left+5,rect.top+5))
                tm.sleep(0.1)
            else:
                break
        if pywinauto.controls.HwndWrapper.HwndWrapper(self.hBuyCmd).IsVisible() == False:
            return {'exchange_order_id':'','error_info':'hCmd is not top valid'}
        if direction == self.od_long:
            return self.buy(symbol, direction, price, volume)
        elif direction == self.od_short:
            return self.sell(symbol, direction, price, volume)
        return ret
        
    def cancel_order(self , exchange_order_id):
        print('CancelOrder:%s'%(exchange_order_id))
        self.check_foreground()
        self.close_all_dialog()         
        if self.check_is_login() == False:
            self.gethandle()

        if self.hTreeLeft == 0 or self.hOrderList == 0:
            return {}
        pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).SetFocus()
        path = "\查询\当日委托"
        pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).EnsureVisible(path)
        rect = pywinauto.controls.common_controls.TreeViewWrapper(self.hTreeLeft).GetItem(path).Rectangle()
        pywinauto.controls.HwndWrapper.HwndWrapper(self.hTreeLeft).ClickInput(coords=(rect.left+5,rect.top+5))
        tm.sleep(0.2)
        l = pywinauto.controls.common_controls.ListViewWrapper(self.hOrderList)
        for i in range(1,10):
            if l.IsVisible():
                break
            else:
                tm.sleep(0.3)
        t = pywinauto.controls.common_controls.ListViewWrapper(self.hOrderList).Texts()
        cc = len(self.lOrderColumn)
        if len(t) > 0 and (len(t)-1)%cc == 0:
            for i in range(1,len(t)):
                r = int((i-1)/cc)
                c = (i-1)%cc
                if self.lOrderColumn[c]['key'] == "exchange_order_id" and t[i] == exchange_order_id :
                    item = l.GetItem(r,c)
                    item.EnsureVisible()
                    item.Click(double=True)
                    hResult = 0
                    for i in range(1,15):
                        try:
                            hResult = pywinauto.findwindows.find_window(class_name='#32770',title_re='提示',process = self.processId)
                            hResultText = win32gui.GetDlgItem(hResult, 0x1B65)
                            hResultConfirm = win32gui.GetDlgItem(hResult, 0x1B67)
                            t = pywinauto.controls.HwndWrapper.HwndWrapper(hResultText).WindowText().splitlines()
                            if len(t) >= 4:
                                if t[0].count('撤单') == 0 or t[3].count(exchange_order_id) == 0 :
                                    return {'error_code':'1','error_info':'撤单确认参数错误'}
                            pywinauto.controls.HwndWrapper.HwndWrapper(hResultConfirm).Click()
                            break
                        except:
                            tm.sleep(0.1)
                    if hResult == 0:
                        return {'error_code':'1','error_info':'撤单确认未正常'}
                    hResult = 0
                    for i in range(1,15):
                        try:
                            hResult = pywinauto.findwindows.find_window(class_name='#32770',title_re='提示',process = self.processId)
                            hResultText = win32gui.GetDlgItem(hResult, 0x1B65)
                            hResultConfirm = win32gui.GetDlgItem(hResult, 0x1B67)
                            t = pywinauto.controls.HwndWrapper.HwndWrapper(hResultText).WindowText().splitlines()
                            ret = {'error_code':'0','error_info':''}
                            print(t)
                            if len(t) == 1 and t[0].count('已提交') > 0:
                                ret['error_info'] = t
                            elif len(t) >= 4:
                                ret['error_code'] = '1'
                                ret['error_info'] = t[3]
                            else:
                                ret['error_code'] = '1'
                                ret['error_info'] = '撤单提交失败'
                            pywinauto.controls.HwndWrapper.HwndWrapper(hResultConfirm).Click()
                            return ret
                            break
                        except:
                            tm.sleep(0.1)
                    if hResult == 0:
                        return {'error_code':'1','error_info':'撤单结果未正常'}
                    break
        return {}
        
    def close_all_dialog(self):
        hdlg = win32gui.FindWindowEx(0, 0, '#32770', None)
        while hdlg > 0:
            hdlg = win32gui.FindWindowEx(0, hdlg, '#32770', None)
            dlgthread, dlgprocessId = win32process.GetWindowThreadProcessId(hdlg)
            if win32gui.IsWindowVisible(hdlg) and dlgprocessId == self.processId and self.threadId == dlgthread:
                win32gui.PostMessage(hdlg, win32con.WM_CLOSE, 0, 0)
        return 0

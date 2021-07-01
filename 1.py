# -*- coding: utf-8 -*-
import win32com.client, pythoncom, sys, time
import pandas as pd
from enum import Enum

XING_PATH = "C:\\eBEST\\xingAPI"

class Server(Enum):
    HTS = 1
    DEMO = 0

class XASessionEventHandler:
    login_state = 0
    def OnLogin(self, code, msg):
        print('on login start')
        if code == "0000":
            print("login succ")
            XASessionEventHandler.login_state = 1
        else:
            XASessionEventHandler.login_state = -1
            print("login fail")

def login(server, id, pwd, cer_pwd, acc, acc_pwd) :
    instXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)
    instXASession.ConnectServer(server, 20001)
    instXASession.Login(id, passwd, cert_passwd, 0, 0)
    while XASessionEventHandler.login_state == 0:
        pythoncom.PumpWaitingMessages()
    login = XASessionEventHandler.login_state
    return login

def wait_for_event(code) :
    while XAQueryEventHandler.query_state == 0:
        pythoncom.PumpWaitingMessages()

    if XAQueryEventHandler.query_code != code :
        print('diff code : wish(', code,')', XAQueryEventHandler.query_code)
        return 0
    XAQueryEventHandler.query_state = 0
    XAQueryEventHandler.query_code = ''
    return 1

class XAQueryEventHandler:
    query_state = 0
    query_code = ''
    T1102_query_state = 0
    T8413_query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandler.query_code = code
        XAQueryEventHandler.query_state = 1

def getFuturesOptions(yyyymm):
    tr_code = 't2301'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = "C:\\eBEST\\xingAPI\\Res\\"+tr_code+".res"
    query.SetFieldData(INBLOCK, "yyyymm", 0, yyyymm) # 월물
    query.SetFieldData(INBLOCK, "gubun", 0, "G") # 구분 mini(M), 정규(G)
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return []
    
    # for futures
    total_data = []
    nCount = query.GetBlockCount(OUTBLOCK)
    for i in range(nCount):        
        gmshCode   = query.GetFieldData(OUTBLOCK, "gmshcode"  , 0).strip()
        histimpv   = query.GetFieldData(OUTBLOCK, "histimpv"  , 0).strip()
        jandatecnt = query.GetFieldData(OUTBLOCK, "jandatecnt", 0).strip()
        cimpv      = query.GetFieldData(OUTBLOCK, "cimpv"     , 0).strip()
        pimpv      = query.GetFieldData(OUTBLOCK, "pimpv"     , 0).strip()
        gmprice    = query.GetFieldData(OUTBLOCK, "gmprice"   , 0).strip()
        gmsign     = query.GetFieldData(OUTBLOCK, "gmsign"    , 0).strip()
        gmchange   = query.GetFieldData(OUTBLOCK, "gmchange"  , 0).strip()
        gmdiff     = query.GetFieldData(OUTBLOCK, "gmdiff"    , 0).strip()
        gmvolume   = query.GetFieldData(OUTBLOCK, "gmvolume"  , 0).strip()
        tmp_data   = [ 
            gmshCode  ,
            histimpv  ,
            jandatecnt,
            cimpv     ,
            pimpv     ,
            gmprice   ,
            gmsign    ,
            gmchange  ,
            gmdiff    ,
            gmvolume ] 
        total_data.append(tmp_data)
    df101 = pd.DataFrame(total_data)
    df101.columns = ['shcode','histimpv','rem','cimpv','pimpv','prx','sgn','chg','updn','vol']

    # for call
    total_data = []
    nCount = query.GetBlockCount(OUTBLOCK1)
    for i in range(nCount) :
        actprice= query.GetFieldData(OUTBLOCK1, "actprice"  , i).strip() #행사가
        optcode = query.GetFieldData(OUTBLOCK1, "optcode"   , i).strip() #종목코드
        price   = query.GetFieldData(OUTBLOCK1, "price"     , i).strip() #행사가
        iv      = query.GetFieldData(OUTBLOCK1, "iv"        , i).strip() #내재변동성
        delta   = query.GetFieldData(OUTBLOCK1, "delt"      , i).strip() #delta
        gamma   = query.GetFieldData(OUTBLOCK1, "gama"      , i).strip() #gamma
        ceta    = query.GetFieldData(OUTBLOCK1, "ceta"      , i).strip() #ceta
        vega    = query.GetFieldData(OUTBLOCK1, "vega"      , i).strip() #vega
        value   = query.GetFieldData(OUTBLOCK1, "value"     , i).strip() #거래대금
        tmp_data= [
            actprice,
            optcode ,
            price   ,
            iv      ,
            delta   ,
            gamma   ,
            ceta    ,
            vega    ,
            value ]
        total_data.append(tmp_data)
    df201 = pd.DataFrame(total_data)
    df201.columns = ['act','code','prx','iv','delt','gama','ceta','vega','value']
    # for put
    total_data = []
    nCount = query.GetBlockCount(OUTBLOCK2)
    for i in range(nCount) :        
        actprice= query.GetFieldData(OUTBLOCK2, "actprice"  , i).strip() #행사가
        optcode = query.GetFieldData(OUTBLOCK2, "optcode"   , i).strip() #종목코드
        price   = query.GetFieldData(OUTBLOCK2, "price"     , i).strip() #행사가
        iv      = query.GetFieldData(OUTBLOCK2, "iv"        , i).strip() #내재변동성
        delta   = query.GetFieldData(OUTBLOCK2, "delt"      , i).strip() #delta
        gamma   = query.GetFieldData(OUTBLOCK2, "gama"      , i).strip() #gamma
        ceta    = query.GetFieldData(OUTBLOCK2, "ceta"      , i).strip() #ceta
        vega    = query.GetFieldData(OUTBLOCK2, "vega"      , i).strip() #vega
        value   = query.GetFieldData(OUTBLOCK2, "value"     , i).strip() #거래대금
        tmp_data= [
            actprice,
            optcode ,
            price   ,
            iv      ,
            delta   ,
            gamma   ,
            ceta    ,
            vega    ,
            value ]
        total_data.append(tmp_data)
    df301 = pd.DataFrame(total_data)
    df301.columns = ['act','code','prx','iv','delt','gama','ceta','vega','value']

    return df101,df201,df301

# 선물/옵션 현재가
def getCurrent(code) :
    tr_code = 't2101'
    INBLOCK = "%sInBlock" % tr_code
    INBLOCK1 = "%sInBlock1" % tr_code
    OUTBLOCK = "%sOutBlock" % tr_code
    OUTBLOCK1 = "%sOutBlock1" % tr_code
    OUTBLOCK2 = "%sOutBlock2" % tr_code
    OUTBLOCK3 = "%sOutBlock3" % tr_code

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
    query.ResFileName = "C:\\eBEST\\xingAPI\\Res\\"+tr_code+".res"
    query.SetFieldData(INBLOCK, "focode", 0, code)
    query.Request(0)

    ret = wait_for_event(tr_code)
    if ret == 0 :
        return 0

    price = query.GetFieldData(OUTBLOCK, "actprice", 0).strip() #가격
    return price

if __name__ == "__main__":    
    RUN_MODE = Server.DEMO
    TODAY = time.strftime("%Y%m%d")
    TODAY_TIME = time.strftime("%H%M%S")
    TODAY_S = time.strftime("%Y-%m-%d")

    if RUN_MODE : #모의투자
        server_add = "hts.ebestsec.co.kr"
        id = "jimsjoo"
        passwd = "sjoo@422"
        cert_passwd = "jimsjoo@3194"
        account_number = "20055436101" 
        account_pwd = "0719"   
    else:
        server_add = "demo.ebestsec.co.kr"
        id = "jimsjoo"    # 본인의 ID로 수정
        passwd = "sjoo@422"
        account_number = '20055436101'
        account_pwd = "0000"   
    
    print('\neBest testing')
    ret = login(server_add, id, passwd, cert_passwd, account_number, account_pwd)
    if ret == 0 :
        print('fail to login')
        quit(0)
    time.sleep(1)

    df1,df2,df3 = getFuturesOptions('202107')

    print('=== df1 ============================='); print(df1)
    print('=== df2 ============================='); print(df2)
    print('=== df3 ============================='); print(df3)
# -*-coding: utf-8 -*-

import win32com.client
import pythoncom
import inspect
import user


class XASessionEvents:
    state = False

    def OnLogin(self, code, msg):
        print("OnLogin : ", code, msg)
        XASessionEvents.state = True

    def OnLogout(self):
        pass

    def OnDisconnect(self):
        pass


class XAQueryEvents:
    state = False

    def OnReceiveData(self, szTrCode):
        print("OnReceiveData : %s" % szTrCode)
        XAQueryEvents.state = True

    def OnReceiveMessage(self, systemError, messageCode, message):
        print("OnReceiveMessage : ", systemError, messageCode, message)


def login(url='hts.ebestsec.co.kr', port=20001, svrtype=0, id='userid', pwd='password', cert='공인인증 비밀번호'):
    session = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
    result = session.ConnectServer(url, port)

    if not result:
        nErrCode = session.GetLastError()
        strErrMsg = session.GetErrorMessage(nErrCode)
        return False, nErrCode, strErrMsg, None, session

    session.Login(id, pwd, cert, svrtype, 0)

    while not XASessionEvents.state:
        pythoncom.PumpWaitingMessages()

    account = []
    n_account = session.GetAccountListCount()

    for i in range(n_account):
        account.append(session.GetAccountList(i))

    return account


def t1101(stockcode):
    instXAQueryT1101 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    instXAQueryT1101.ResFileName = "C:\\eBEST\\xingAPI\\Res\\t1101.res"
    instXAQueryT1101.SetFieldData("t1101InBlock", "shcode", 0, stockcode)
    instXAQueryT1101.Request(0)

    while not XAQueryEvents.state:
        pythoncom.PumpWaitingMessages()
    XAQueryEvents.state = False

    data = dict()
    data['code'] = instXAQueryT1101.GetFieldData("t1101OutBlock", "shcode", 0)
    data['name'] = instXAQueryT1101.GetFieldData("t1101OutBlock", "hname", 0)
    data['price'] = instXAQueryT1101.GetFieldData("t1101OutBlock", "price", 0)
    data['ho_status'] = instXAQueryT1101.GetFieldData("t1101OutBlock", "ho_status", 0)  # 동시구분 1:장중, 2:시간외, 3:장전/장중/장마감 동시

    print(data)
    return data


def CSPAT00600(input_data):
    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    myname = inspect.currentframe().f_code.co_name
    inblock1 = "%sInBlock1" % myname

    query.LoadFromResFile("C:\\eBEST\\xingAPI\\Res\\CSPAT00600.res")
    query.SetFieldData(inblock1, "AcntNo", 0, input_data['계좌번호'])
    query.SetFieldData(inblock1, "InptPwd", 0, input_data['입력비밀번호'])
    query.SetFieldData(inblock1, "IsuNo", 0, input_data['종목번호'])
    query.SetFieldData(inblock1, "OrdQty", 0, input_data['주문수량'])
    query.SetFieldData(inblock1, "OrdPrc", 0, input_data['주문가'])
    query.SetFieldData(inblock1, "BnsTpCode", 0, input_data['매매구분'])
    query.SetFieldData(inblock1, "OrdprcPtnCode", 0, input_data['호가유형코드'])
    query.SetFieldData(inblock1, "MgntrnCode", 0, input_data['신용거래코드'])
    query.SetFieldData(inblock1, "LoanDt", 0, input_data['대출일'])
    query.SetFieldData(inblock1, "OrdCndiTpCode", 0, input_data['주문조건구문'])
    query.Request(0)

    while not XAQueryEvents.state:
        pythoncom.PumpWaitingMessages()
    XAQueryEvents.state = False


def buy_order(_code):
    _data = t1101(_code)
    if _data['ho_status'] == '3':  # 동시호가 시간에서만 주문 실행
        _amount = divmod(trade_price, int(_data['price']))[0]
        order_input = {
            '계좌번호': acc[0],  # 계좌가 여러개일 경우 수정필요
            '입력비밀번호': user.trade_pwd,
            '종목번호': _code,
            '주문수량': _amount,
            '주문가': 0,
            '매매구분': '2',  # 매수
            '호가유형코드': '03',  # 시장가
            '신용거래코드': '000',
            '대출일': '',
            '주문조건구문': '0'
        }
        CSPAT00600(order_input)


if __name__ == '__main__':
    acc = login(id=user.id, pwd=user.pwd, cert=user.cert_pwd)
    print(acc)  # 계좌번호

    codes = ['233740', '233160', '229200', '232080']
    trade_price = 20000

    for code in codes:
        buy_order(code)

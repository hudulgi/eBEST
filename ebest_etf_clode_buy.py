# -*-coding: utf-8 -*-

import win32com.client
import pythoncom
import inspect
import user
import pandas as pd


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
    myname = inspect.currentframe().f_code.co_name
    inblock = "%sInBlock" % myname
    outblock = "%sOutBlock" % myname
    resfile = "%s\\Res\\%s.res" % (resdir, myname)

    instXAQueryT1101 = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    instXAQueryT1101.ResFileName = resfile
    instXAQueryT1101.SetFieldData(inblock, "shcode", 0, stockcode)
    instXAQueryT1101.Request(0)

    while not XAQueryEvents.state:
        pythoncom.PumpWaitingMessages()
    XAQueryEvents.state = False

    data = dict()
    data['code'] = instXAQueryT1101.GetFieldData(outblock, "shcode", 0)
    data['name'] = instXAQueryT1101.GetFieldData(outblock, "hname", 0)
    data['price'] = instXAQueryT1101.GetFieldData(outblock, "price", 0)
    data['ho_status'] = instXAQueryT1101.GetFieldData(outblock, "ho_status", 0)  # 동시구분 1:장중, 2:시간외, 3:장전/장중/장마감 동시

    print(data)
    return data


def CSPAT00600(input_data):
    """
    현물주문
    """
    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    myname = inspect.currentframe().f_code.co_name
    inblock1 = "%sInBlock1" % myname
    resfile = "%s\\Res\\%s.res" % (resdir, myname)

    query.LoadFromResFile(resfile)
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


def t0424(accno='', passwd='', prcgb='1', chegb='0', dangb='0', charge='1', cts_expcode=''):
    '''
    주식잔고2
    '''
    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)

    myname = inspect.currentframe().f_code.co_name
    inblock = "%sInBlock" % myname
    outblock = "%sOutBlock" % myname
    outblock1 = "%sOutBlock1" % myname
    resfile = "%s\\Res\\%s.res" % (resdir, myname)

    query.LoadFromResFile(resfile)
    query.SetFieldData(inblock, "accno", 0, accno)
    query.SetFieldData(inblock, "passwd", 0, passwd)
    query.SetFieldData(inblock, "prcgb", 0, prcgb)
    query.SetFieldData(inblock, "chegb", 0, chegb)
    query.SetFieldData(inblock, "dangb", 0, dangb)
    query.SetFieldData(inblock, "charge", 0, charge)
    query.SetFieldData(inblock, "cts_expcode", 0, cts_expcode)
    query.Request(0)

    while not XAQueryEvents.state:
        pythoncom.PumpWaitingMessages()

    result = []
    nCount = query.GetBlockCount(outblock)
    for i in range(nCount):
        sunamt = int(query.GetFieldData(outblock, "sunamt", i).strip())  # 추정순자산
        dtsunik = int(query.GetFieldData(outblock, "dtsunik", i).strip())  # 실현손익
        mamt = int(query.GetFieldData(outblock, "mamt", i).strip())  # 매입금액
        sunamt1 = int(query.GetFieldData(outblock, "sunamt1", i).strip())  # 추정D2예수금
        cts_expcode = query.GetFieldData(outblock, "cts_expcode", i).strip()  # CTS_종목번호
        tappamt = int(query.GetFieldData(outblock, "tappamt", i).strip())  # 평가금액
        tdtsunik = int(query.GetFieldData(outblock, "tdtsunik", i).strip())  # 평가손익

        lst = [sunamt, dtsunik, mamt, sunamt1, cts_expcode, tappamt, tdtsunik]
        result.append(lst)

    columns = ['추정순자산', '실현손익', '매입금액', '추정D2예수금', 'CTS_종목번호', '평가금액', '평가손익']
    df_outblock = pd.DataFrame(data=result, columns=columns)

    result = []
    nCount = query.GetBlockCount(outblock1)
    for i in range(nCount):
        expcode = query.GetFieldData(outblock1, "expcode", i).strip()  # 종목번호
        jangb = query.GetFieldData(outblock1, "jangb", i).strip()  # 잔고구분
        janqty = int(query.GetFieldData(outblock1, "janqty", i).strip())  # 잔고수량
        mdposqt = int(query.GetFieldData(outblock1, "mdposqt", i).strip())  # 매도가능수량
        pamt = int(query.GetFieldData(outblock1, "pamt", i).strip())  # 평균단가
        mamt = int(query.GetFieldData(outblock1, "mamt", i).strip())  # 매입금액
        sinamt = int(query.GetFieldData(outblock1, "sinamt", i).strip())  # 대출금액
        lastdt = query.GetFieldData(outblock1, "lastdt", i).strip()  # 만기일자
        msat = int(query.GetFieldData(outblock1, "msat", i).strip())  # 당일매수금액
        mpms = int(query.GetFieldData(outblock1, "mpms", i).strip())  # 당일매수단가
        mdat = int(query.GetFieldData(outblock1, "mdat", i).strip())  # 당일매도금액
        mpmd = int(query.GetFieldData(outblock1, "mpmd", i).strip())  # 당일매도단가
        jsat = int(query.GetFieldData(outblock1, "jsat", i).strip())  # 전일매수금액
        jpms = int(query.GetFieldData(outblock1, "jpms", i).strip())  # 전일매수단가
        jdat = int(query.GetFieldData(outblock1, "jdat", i).strip())  # 전일매도금액
        jpmd = int(query.GetFieldData(outblock1, "jpmd", i).strip())  # 전일매도단가
        sysprocseq = int(query.GetFieldData(outblock1, "sysprocseq", i).strip())  # 처리순번
        loandt = query.GetFieldData(outblock1, "loandt", i).strip()  # 대출일자
        hname = query.GetFieldData(outblock1, "hname", i).strip()  # 종목명
        marketgb = query.GetFieldData(outblock1, "marketgb", i).strip()  # 시장구분
        jonggb = query.GetFieldData(outblock1, "jonggb", i).strip()  # 종목구분
        janrt = float(query.GetFieldData(outblock1, "janrt", i).strip())  # 보유비중
        price = int(query.GetFieldData(outblock1, "price", i).strip())  # 현재가
        appamt = int(query.GetFieldData(outblock1, "appamt", i).strip())  # 평가금액
        dtsunik = int(query.GetFieldData(outblock1, "dtsunik", i).strip())  # 평가손익
        sunikrt = float(query.GetFieldData(outblock1, "sunikrt", i).strip())  # 수익율
        fee = int(query.GetFieldData(outblock1, "fee", i).strip())  # 수수료
        tax = int(query.GetFieldData(outblock1, "tax", i).strip())  # 제세금
        sininter = int(query.GetFieldData(outblock1, "sininter", i).strip())  # 신용이자

        lst = [expcode, jangb, janqty, mdposqt, pamt, mamt, sinamt, lastdt, msat,
               mpms, mdat, mpmd, jsat, jpms, jdat, jpmd,
               sysprocseq, loandt, hname, marketgb, jonggb, janrt, price, appamt, dtsunik, sunikrt, fee, tax, sininter]
        result.append(lst)

    columns = ['종목번호', '잔고구분', '잔고수량', '매도가능수량', '평균단가', '매입금액', '대출금액', '만기일자', '당일매수금액', ' 당일매수단가', '당일매도금액',
               '당일매도단가', '전일매수금액', '전일매수단가', '전일매도금액', '전일매도단가', ' 처리순번', '대출일자', '종목명', '시장구분', '종목구분', '보유비중', '현재가',
               '평가금액', '평가손익', '수익율', '수수료', '제세금', '신용이자']
    df_outblock1 = pd.DataFrame(data=result, columns=columns)

    XAQueryEvents.state = False

    return df_outblock, df_outblock1


def buy_order(_code):
    _data = t1101(_code)
    if _data['ho_status'] == '3':  # 동시호가 시간에서만 주문 실행
        _amount = divmod(trade_price, int(_data['price']))[0]  # 주문 수량 계산
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
        CSPAT00600(order_input)  # 주문 실행


def sell_order_all(_acc, _pwd):
    """
    잔고 시장가 일괄매도
    """
    jango1, jango2 = t0424(accno=_acc, passwd=_pwd)

    for i in range(jango2.shape[0]):
        temp = jango2.iloc[i]
        if temp['매도가능수량'] > 0:
            _code = temp['종목번호']
            _name = temp['종목명']
            _amount = temp['매도가능수량']
            print(_code, _name, _amount)
            order_input = {
                '계좌번호': acc[0],  # 계좌가 여러개일 경우 수정필요
                '입력비밀번호': user.trade_pwd,
                '종목번호': _code,
                '주문수량': _amount,
                '주문가': 0,
                '매매구분': '1',  # 매도
                '호가유형코드': '03',  # 시장가
                '신용거래코드': '000',
                '대출일': '',
                '주문조건구문': '0'
            }
            CSPAT00600(order_input)


if __name__ == '__main__':
    resdir = 'C:\\eBEST\\xingAPI'

    acc = login(id=user.id, pwd=user.pwd, cert=user.cert_pwd)
    print(acc)  # 계좌번호

    codes = ['233740', '233160', '229200', '232080']  # KODEX, TIGER 코스닥150 & 레버리지
    trade_price = 20000

    for code in codes:
        buy_order(code)

    # sell_order_all(acc[0], user.trade_pwd)
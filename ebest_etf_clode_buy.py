# -*-coding: utf-8 -*-

import win32com.client
import pythoncom
import os
import sys
import inspect
import pandas as pd
from pandas import DataFrame, Series, Panel
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

    name = instXAQueryT1101.GetFieldData("t1101OutBlock", "hname", 0)
    price = instXAQueryT1101.GetFieldData("t1101OutBlock", "price", 0)
    print(name)
    print(price)


def CSPAT00600(계좌번호, 입력비밀번호, 종목번호, 주문수량, 주문가, 매매구분, 호가유형코드, 신용거래코드, 주문조건구분):
    pathname = os.path.dirname(sys.argv[0])
    resdir = os.path.abspath(pathname)

    query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)

    MYNAME = inspect.currentframe().f_code.co_name
    INBLOCK = "%sInBlock" % MYNAME
    INBLOCK1 = "%sInBlock1" % MYNAME
    OUTBLOCK = "%sOutBlock" % MYNAME
    OUTBLOCK1 = "%sOutBlock1" % MYNAME
    OUTBLOCK2 = "%sOutBlock2" % MYNAME
    RESFILE = "%s\\Res\\%s.res" % (resdir, MYNAME)

    print(MYNAME, RESFILE)

    query.LoadFromResFile(RESFILE)
    query.SetFieldData(INBLOCK1, "AcntNo", 0, 계좌번호)
    query.SetFieldData(INBLOCK1, "InptPwd", 0, 입력비밀번호)
    query.SetFieldData(INBLOCK1, "IsuNo", 0, 종목번호)
    query.SetFieldData(INBLOCK1, "OrdQty", 0, 주문수량)
    query.SetFieldData(INBLOCK1, "OrdPrc", 0, 주문가)
    query.SetFieldData(INBLOCK1, "BnsTpCode", 0, 매매구분)
    query.SetFieldData(INBLOCK1, "OrdprcPtnCode", 0, 호가유형코드)
    query.SetFieldData(INBLOCK1, "MgntrnCode", 0, 신용거래코드)
    # query.SetFieldData(INBLOCK1, "LoanDt", 0, 대출일)
    query.SetFieldData(INBLOCK1, "OrdCndiTpCode", 0, 주문조건구분)
    query.Request(0)

    while XAQueryEvents.상태 == False:
        pythoncom.PumpWaitingMessages()

    result = []
    nCount = query.GetBlockCount(OUTBLOCK1)
    for i in range(nCount):
        레코드갯수 = int(query.GetFieldData(OUTBLOCK1, "RecCnt", i).strip())
        계좌번호 = query.GetFieldData(OUTBLOCK1, "AcntNo", i).strip()
        입력비밀번호 = query.GetFieldData(OUTBLOCK1, "InptPwd", i).strip()
        종목번호 = query.GetFieldData(OUTBLOCK1, "IsuNo", i).strip()
        주문수량 = int(query.GetFieldData(OUTBLOCK1, "OrdQty", i).strip())
        주문가 = query.GetFieldData(OUTBLOCK1, "OrdPrc", i).strip()
        매매구분 = query.GetFieldData(OUTBLOCK1, "BnsTpCode", i).strip()
        호가유형코드 = query.GetFieldData(OUTBLOCK1, "OrdprcPtnCode", i).strip()
        프로그램호가유형코드 = query.GetFieldData(OUTBLOCK1, "PrgmOrdprcPtnCode", i).strip()
        공매도가능여부 = query.GetFieldData(OUTBLOCK1, "StslAbleYn", i).strip()
        공매도호가구분 = query.GetFieldData(OUTBLOCK1, "StslOrdprcTpCode", i).strip()
        통신매체코드 = query.GetFieldData(OUTBLOCK1, "CommdaCode", i).strip()
        신용거래코드 = query.GetFieldData(OUTBLOCK1, "MgntrnCode", i).strip()
        대출일 = query.GetFieldData(OUTBLOCK1, "LoanDt", i).strip()
        회원번호 = query.GetFieldData(OUTBLOCK1, "MbrNo", i).strip()
        주문조건구분 = query.GetFieldData(OUTBLOCK1, "OrdCndiTpCode", i).strip()
        전략코드 = query.GetFieldData(OUTBLOCK1, "StrtgCode", i).strip()
        그룹ID = query.GetFieldData(OUTBLOCK1, "GrpId", i).strip()
        주문회차 = int(query.GetFieldData(OUTBLOCK1, "OrdSeqNo", i).strip())
        포트폴리오번호 = int(query.GetFieldData(OUTBLOCK1, "PtflNo", i).strip())
        바스켓번호 = int(query.GetFieldData(OUTBLOCK1, "BskNo", i).strip())
        트렌치번호 = int(query.GetFieldData(OUTBLOCK1, "TrchNo", i).strip())
        아이템번호 = int(query.GetFieldData(OUTBLOCK1, "ItemNo", i).strip())
        운용지시번호 = query.GetFieldData(OUTBLOCK1, "OpDrtnNo", i).strip()
        유동성공급자여부 = query.GetFieldData(OUTBLOCK1, "LpYn", i).strip()
        반대매매구분 = query.GetFieldData(OUTBLOCK1, "CvrgTpCode", i).strip()

        lst = [레코드갯수, 계좌번호, 입력비밀번호, 종목번호, 주문수량, 주문가, 매매구분, 호가유형코드, 프로그램호가유형코드, 공매도가능여부, 공매도호가구분, 통신매체코드, 신용거래코드, 대출일,
               회원번호, 주문조건구분, 전략코드, 그룹ID, 주문회차, 포트폴리오번호, 바스켓번호, 트렌치번호, 아이템번호, 운용지시번호, 유동성공급자여부, 반대매매구분]
        result.append(lst)

    columns = ['레코드갯수', '계좌번호', '입력비밀번호', '종목번호', '주문수량', '주문가', '매매구분', '호가유형코드', '프로그램호가유형코드', '공매도가능여부', '공매도호가구분',
               '통신매체코드', '신용거래코드', '대출일', '회원번호', '주문조건구분', '전략코드', '그룹ID', '주문회차', '포트폴리오번호', '바스켓번호', '트렌치번호',
               '아이템번호', '운용지시번호', '유동성공급자여부', '반대매매구분']
    df = DataFrame(data=result, columns=columns)

    result = []
    nCount = query.GetBlockCount(OUTBLOCK2)
    for i in range(nCount):
        레코드갯수 = int(query.GetFieldData(OUTBLOCK2, "RecCnt", i).strip())
        주문번호 = int(query.GetFieldData(OUTBLOCK2, "OrdNo", i).strip())
        주문시각 = query.GetFieldData(OUTBLOCK2, "OrdTime", i).strip()
        주문시장코드 = query.GetFieldData(OUTBLOCK2, "OrdMktCode", i).strip()
        주문유형코드 = query.GetFieldData(OUTBLOCK2, "OrdPtnCode", i).strip()
        단축종목번호 = query.GetFieldData(OUTBLOCK2, "ShtnIsuNo", i).strip()
        관리사원번호 = query.GetFieldData(OUTBLOCK2, "MgempNo", i).strip()
        주문금액 = int(query.GetFieldData(OUTBLOCK2, "OrdAmt", i).strip())
        예비주문번호 = int(query.GetFieldData(OUTBLOCK2, "SpareOrdNo", i).strip())
        반대매매일련번호 = int(query.GetFieldData(OUTBLOCK2, "CvrgSeqno", i).strip())
        예약주문번호 = int(query.GetFieldData(OUTBLOCK2, "RsvOrdNo", i).strip())
        실물주문수량 = int(query.GetFieldData(OUTBLOCK2, "SpotOrdQty", i).strip())
        재사용주문수량 = int(query.GetFieldData(OUTBLOCK2, "RuseOrdQty", i).strip())
        현금주문금액 = int(query.GetFieldData(OUTBLOCK2, "MnyOrdAmt", i).strip())
        대용주문금액 = int(query.GetFieldData(OUTBLOCK2, "SubstOrdAmt", i).strip())
        재사용주문금액 = int(query.GetFieldData(OUTBLOCK2, "RuseOrdAmt", i).strip())
        계좌명 = query.GetFieldData(OUTBLOCK2, "AcntNm", i).strip()
        종목명 = query.GetFieldData(OUTBLOCK2, "IsuNm", i).strip()

        lst = [레코드갯수, 주문번호, 주문시각, 주문시장코드, 주문유형코드, 단축종목번호, 관리사원번호, 주문금액, 예비주문번호, 반대매매일련번호, 예약주문번호, 실물주문수량, 재사용주문수량,
               현금주문금액, 대용주문금액, 재사용주문금액, 계좌명, 종목명]
        result.append(lst)

    columns = ['레코드갯수', '주문번호', '주문시각', '주문시장코드', '주문유형코드', '단축종목번호', '관리사원번호', '주문금액', '예비주문번호', '반대매매일련번호', '예약주문번호',
               '실물주문수량', '재사용주문수량', '현금주문금액', '대용주문금액', '재사용주문금액', '계좌명', '종목명']
    df1 = DataFrame(data=result, columns=columns)

    XAQueryEvents.상태 = False

    return (df, df1)


if __name__ == '__main__':
    acc = login(id=user.id, pwd=user.pwd, cert=user.cert_pwd)
    codes = ['233740', '233160', '229200', '232080']
    for code in codes:
        t1101(code)

import sys
from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMainWindow


class Main(QMainWindow):

    def __init__(self):
        super().__init__()

        # 일반 TR OCX
        self.IndiTR = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")
        self.IndiReal = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")

        # TR 관련 변수
        self.IndiTR.ReceiveData.connect(self.ReceiveData)
        self.IndiTR.ReceiveSysMsg.connect(self.ReceiveSysMsg)
        self.rqidList = {}

        # Real 관련 변수
        self.IndiReal.ReceiveRTData.connect(self.ReceiveRTData)

        # 신한i Indi 자동로그인
        while True:
            login = self.IndiTR.StartIndi('ID',
                                          'PASS',
                                          'CERT-PASS',
                                          'C:/SHINHAN-i/indi/giexpertstarter.exe')
            print(login)
            if login:
                break

    def ReceiveData(self, rqid):
        if self.rqidList[rqid] == "stock_mst":
            self.stock_mst()
        if self.rqidList[rqid] == "TR_SCHART":
            self.TR_SCAHRT()
        if self.rqidList[rqid] == "SC":
            self.SC()
        if self.rqidList[rqid] == "AccountList":
            self.AccountList()
        if self.rqidList[rqid] == "SABA200QB":
            self.AccountLookUp()
        if self.rqidList[rqid] == "SABA101U1":
            self.order()

        self.rqidList.__delitem__(rqid)

    def ReceiveRTData(self, RealType):
        if RealType == "SC":
            DATA = {}
            DATA['ISIN_CODE'] = self.IndiReal.dynamicCall("GetSingleData(int)", 0)  # 표준코드
            DATA['CODE'] = self.IndiReal.dynamicCall("GetSingleData(int)", 1)  # 단축코드
            DATA['Time'] = self.IndiReal.dynamicCall("GetSingleData(int)", 2)  # 채결시간
            DATA['Close'] = self.IndiReal.dynamicCall("GetSingleData(int)", 3)  # 현재가
            DATA['Vol'] = self.IndiReal.dynamicCall("GetSingleData(int)", 7)  # 누적거래량
            DATA['TRADING_VALUE'] = self.IndiReal.dynamicCall("GetSingleData(int)", 8)  # 누적거래대금
            DATA['ContQty'] = self.IndiReal.dynamicCall("GetSingleData(int)", 9)  # 단위채결량
            DATA['Open'] = self.IndiReal.dynamicCall("GetSingleData(int)", 10)  # 시가
            DATA['High'] = self.IndiReal.dynamicCall("GetSingleData(int)", 11)  # 고가
            DATA['Low'] = self.IndiReal.dynamicCall("GetSingleData(int)", 12)  # 저가
            print(DATA)

    def ReceiveSysMsg(self, MsgID):
        print('System Message Received = ', MsgID)

    # =================================================================================

    def req_stock_mst(self):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "stock_mst")
        rqid = self.IndiTR.dynamicCall("RequestData()")
        self.rqidList[rqid] = "stock_mst"

    def req_TR_SCHART(self, stock_code, candleUnit, term, startTerm, endTerm):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "TR_SCHART")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, stock_code)  # 단축코드
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 1, candleUnit)  # 1:분봉, D:일봉, W:주봉, M:월봉
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 2, term)  # 분봉: 1~30, 일/주/월 : 1
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 3, startTerm)  # 시작일(YYYYMMDD)
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 4, endTerm)  # 종료일(YYYYMMDD)
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 5, "9999")  # 조회갯수(1~9999)
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청

        self.rqidList[rqid] = "TR_SCHART"

    def req_SC(self, stockCode):
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SB", stockCode)
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SC", stockCode)
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SH", stockCode)

        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "SC")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, stockCode)  # 인풋 : 단축코드

        rqid = self.IndiTR.dynamicCall("RequestData()")  # RequestRTReg()를 사용해 TR을 등록
        print('RQID :', rqid)

        self.rqidList[rqid] = "SC"

        print(self.rqidList)

    def req_AccountList(self):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "AccountList")
        rqid = self.IndiTR.dynamicCall("RequestData()")
        self.rqidList[rqid] = "AccountList"

    def req_AccountLookUp(self, AccountNum, AccountPass):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "SABA200QB")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, AccountNum)  # 계좌번호
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 1, "01")  # 상품구분 항상 '01'
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 2, AccountPass)  # 비밀번호
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청
        self.rqidList[rqid] = "SABA200QB"

    def req_order(self, account_num, pwd, code, price, count, order_type, call_type):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "SABA101U1")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, str(account_num))  # 계좌번호
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 1, "01")  # 상품구분
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 2, str(pwd))  # 비밀번호
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 5, '0')  # 선물대용매도구분
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 6, '00')  # 신용거래구분
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 7, str(order_type))  # 매수매도 구분
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 8, 'A' + str(code))  # 종목코드
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 9, str(count))  # 주문수량
        #ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 10, str(price))  # 주문가격
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 11, '1')  # 정규장
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 12, call_type)  # 호가유형, 1: 시장가, X:최유리, Y:최우선
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 13, '0')  # 주문조건, 0:일반, 3:IOC, 4:FOK
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 14, '0')  # 신용대출
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 21, 'Y')  # 결과 출력 여부
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청

        # 요청한 ID를 저장합니다.
        self.rqidList[rqid] = "SABA101U1"
        print("매매TR요청 : ", rqid)

    # =================================================================================

    def stock_mst(self):
        count = self.IndiTR.dynamicCall("GetMultiRowCount()")
        result = []

        for i in range(0, count):
            DATA = {}

            DATA['ISIN_CODE'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 0)  # 표준코드
            DATA['CODE'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 1)  # 단축코드
            DATA['MARKET'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 2)  # 장구분
            DATA['NAME'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 3)  # 종목명
            DATA['SECTOR'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 4)  # KOSPI200 세부업종
            DATA['SETTLEMENT'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 5)  # 결산연월
            DATA['SEC'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 6)  # 거래정지구분
            DATA['MANAGEMENT'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 7)  # 관리구분
            DATA['ALERT'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 8)  # 시장경보구분코드
            DATA['ROCK'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 9)  # 락구분
            DATA['INVALID'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 10)  # 불성실공시지정여부
            DATA['MARGIN'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 11)  # 증거금 구분
            DATA['CREDIT'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 12)  # 신용증거금 구분
            DATA['ETF'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 13)  # ETF 구분
            DATA['PART'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 14)  # 소속구분

            result.append(DATA)

    def TR_SCAHRT(self):
        count = self.IndiTR.dynamicCall("GetMultiRowCount()")
        result = []

        for i in range(0, count):
            DATA = {}
            DATA['DATE'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 0)  # 일자
            DATA['TIME'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 1)  # 시간
            DATA['OPEN'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 2)  # 시가
            DATA['HIGH'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 3)  # 고가
            DATA['Low'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 4)  # 저가
            DATA['Close'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 5)  # 종가
            DATA['Price_ADJ'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 6)  # 주가수정계수
            DATA['Vol_ADJ'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 7)  # 거래량 수정계수
            DATA['Rock'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 8)  # 락구분
            DATA['Vol'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 9)  # 단위거래량
            DATA['Trading_Value'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 10)  # 단위거래대금

            result.append(DATA)

    def SC(self):
        DATA = {}
        DATA['ISIN_CODE'] = self.IndiTR.dynamicCall("GetSingleData(int)", 0)  # 표준코드
        DATA['CODE'] = self.IndiTR.dynamicCall("GetSingleData(int)", 1)  # 단축코드
        DATA['Time'] = self.IndiTR.dynamicCall("GetSingleData(int)", 2)  # 채결시간
        DATA['Close'] = self.IndiTR.dynamicCall("GetSingleData(int)", 3)  # 현재가
        DATA['Vol'] = self.IndiTR.dynamicCall("GetSingleData(int)", 7)  # 누적거래량
        DATA['TRADING_VALUE'] = self.IndiTR.dynamicCall("GetSingleData(int)", 8)  # 누적거래대금
        DATA['ContQty'] = self.IndiTR.dynamicCall("GetSingleData(int)", 9)  # 단위채결량
        DATA['Open'] = self.IndiTR.dynamicCall("GetSingleData(int)", 10)  # 시가
        DATA['High'] = self.IndiTR.dynamicCall("GetSingleData(int)", 11)  # 고가
        DATA['Low'] = self.IndiTR.dynamicCall("GetSingleData(int)", 12)  # 저가

        # 실시간 등록
        ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "SC", DATA['CODE'])
        print(ret, DATA)

    def AccountList(self):
        count = self.IndiTR.dynamicCall("GetMultiRowCount()")
        result = []

        for i in range(0, count):
            DATA = {}
            DATA['CODE'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 0)  # 단축코드
            DATA['NAME'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 1)  # 종목명
            result.append(DATA)
            print(DATA['CODE'], DATA['NAME'])

    def AccountLookUp(self):
        nCnt = self.IndiTR.dynamicCall("GetMultiRowCount()")
        result = []

        for i in range(0, nCnt):
            DATA = {}
            DATA['ISIN_CODE'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 0)  # 표준코드
            DATA['NAME'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 1)  # 종목명
            DATA['NUM'] = int(self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 2))  # 결재일 잔고 수량
            DATA['SELL_UNFINISHED_NUM'] = int(self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 3))  # 매도 미체결 수량
            DATA['BUY_UNFINISHED_NUM'] = int(self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 4))  # 매수 미체결 수량
            DATA['CURRENT_PRC'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 5)  # 현재가
            DATA['AVG_PRC'] = float(self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 6))  # 평균단가
            DATA['CREDIT_NUM'] = int(self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 7))  # 신용잔고수량
            DATA['KOSPI_NUM'] = int(self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 8))  # 코스피대용수량
            result.append(DATA)
            print(DATA)

    def order(self):
        DATA = {}
        DATA['Order_Num'] = self.IndiTR.dynamicCall("GetSingleData(int)", 0)  # 주문번호
        DATA['Num'] = self.IndiTR.dynamicCall("GetSingleData(int)", 2)  # 메시지 구분
        DATA['Msg1'] = self.IndiTR.dynamicCall("GetSingleData(int)", 3)  # 메시지1
        DATA['Msg2'] = self.IndiTR.dynamicCall("GetSingleData(int)", 4)  # 메시지2
        DATA['Msg3'] = self.IndiTR.dynamicCall("GetSingleData(int)", 5)  # 메시지3
        print("매수 및 매도 주문결과 :", DATA)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    Indi = Main()
    app.exec_()

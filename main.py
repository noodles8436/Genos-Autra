import sys

from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMainWindow

from Analayzer import Analayzer


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
        self.hogaList = {}

        # Real 관련 변수
        self.IndiReal.ReceiveRTData.connect(self.ReceiveRTData)

        # Analyzer
        self.Analyzer = Analayzer()

        # 신한i Indi 자동로그인
        # TODO : 입력 받기
        while True:
            login = self.IndiTR.StartIndi('xodnr8436',
                                          '!noodles5169',
                                          '!noodles8436',
                                          'C:/SHINHAN-i/indi/giexpertstarter.exe')
            print(login)
            if login:
                break

        # TODO : 입력 받기
        self.AccountID = "27038563451"
        self.AccountPass = "1215"

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

        if self.rqidList[rqid] == "SH":
            self.SH(self.hogaList[rqid])
            self.hogaList.__delitem__(rqid)

        if self.rqidList[rqid] == "Rebalancing":
            self.rebalancing()

        self.rqidList.__delitem__(rqid)

    def ReceiveRTData(self, RealType):
        if RealType == "SC":
            DATA = {}
            DATA['ISIN_CODE'] = self.IndiReal.dynamicCall("GetSingleData(int)", 0)  # 표준코드
            DATA['CODE'] = self.IndiReal.dynamicCall("GetSingleData(int)", 1)  # 단축코드
            DATA['Time'] = self.IndiReal.dynamicCall("GetSingleData(int)", 2)  # 채결시간
            DATA['Close'] = self.IndiReal.dynamicCall("GetSingleData(int)", 3)  # 현재가
            DATA['DCY'] = self.IndiReal.dynamicCall("GetSingleData(int)", 4)  # 전일대비구분
            DATA['CY'] = self.IndiReal.dynamicCall("GetSingleData(int)", 5)  # 전일대비
            DATA['RCY'] = self.IndiReal.dynamicCall("GetSingleData(int)", 6)  # 전일대비율
            DATA['Vol'] = self.IndiReal.dynamicCall("GetSingleData(int)", 7)  # 누적거래량
            DATA['TRADING_VALUE'] = self.IndiReal.dynamicCall("GetSingleData(int)", 8)  # 누적거래대금
            DATA['ContQty'] = self.IndiReal.dynamicCall("GetSingleData(int)", 9)  # 단위채결량
            DATA['Open'] = self.IndiReal.dynamicCall("GetSingleData(int)", 10)  # 시가
            DATA['High'] = self.IndiReal.dynamicCall("GetSingleData(int)", 11)  # 고가
            DATA['Low'] = self.IndiReal.dynamicCall("GetSingleData(int)", 12)  # 저가
            DATA['POT'] = self.IndiReal.dynamicCall("GetSingleData(int)", 22)  # 거래강도
            DATA['COP'] = self.IndiReal.dynamicCall("GetSingleData(int)", 24)  # 체결강도
            DATA['DOC'] = self.IndiReal.dynamicCall("GetSingleData(int)", 25)  # 체결 매도매수 구분

            if self.Analyzer.check(DATA):
                self.req_TR_SH(DATA['CODE'], DATA['Time'])

    def ReceiveSysMsg(self, MsgID):
        print('System Message Received = ', MsgID)
        print('Get Error Message : ', self.IndiTR.GetErrorMessage())

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

    def req_TR_SH(self, stockCode, stockTime):
        self.IndiTR.dynamicCall("SetQueryName(QString)", "SH")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, stockCode)  # 단축코드
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청

        self.rqidList[rqid] = "SH"
        self.hogaList[rqid] = stockTime

    def req_SC(self, stockCode):
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SB", stockCode)
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SC", stockCode)
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SH", stockCode)

        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "SC")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, stockCode)  # 인풋 : 단축코드

        rqid = self.IndiTR.dynamicCall("RequestData()")  # RequestRTReg()를 사용해 TR을 등록

        self.rqidList[rqid] = "SC"
        print(' NOW -> ', self.rqidList)

    def req_Real_SC(self, stockCode):
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SB", stockCode)
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SC", stockCode)
        self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "SH", stockCode)
        ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "SC", stockCode)

    def req_AccountList(self):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "AccountList")
        rqid = self.IndiTR.dynamicCall("RequestData()")
        self.rqidList[rqid] = "AccountList"

    def req_AccountLookUp(self):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "SABA200QB")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, self.AccountID)  # 계좌번호
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 1, "01")  # 상품구분 항상 '01'
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 2, self.AccountPass)  # 비밀번호
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청
        self.rqidList[rqid] = "SABA200QB"

    # TODO : 경우에 따른 안자값 불필요 제거 - 오버로딩 하던가?
    def req_order(self, code, price, count, order_type, call_type):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "SABA101U1")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, self.AccountID)  # 계좌번호
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 1, "01")  # 상품구분
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 2, self.AccountPass)  # 비밀번호
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 5, '0')  # 선물대용매도구분
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 6, '00')  # 신용거래구분
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 7, str(order_type))  # 매수매도 구분
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 8, 'A' + str(code))  # 종목코드
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 9, str(count))  # 주문수량
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 10, str(price))  # 주문가격
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 11, '1')  # 정규장
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 12, call_type)  # 호가유형, 1: 시장가, X:최유리, Y:최우선
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 13, '0')  # 주문조건, 0:일반, 3:IOC, 4:FOK
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 14, '0')  # 신용대출
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 21, 'Y')  # 결과 출력 여부
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청

        self.rqidList[rqid] = "SABA101U1"
        print("매매TR요청 : ", rqid)

    def req_rebalancing(self):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "SABA200QB")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, self.AccountID)  # 계좌번호
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 1, "01")  # 상품구분 항상 '01'
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 2, self.AccountPass)  # 비밀번호
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청
        self.rqidList[rqid] = "Rebalancing"

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

        self.Analyzer.setTotalStockList(result)
        print('모든 종목 받고 SC 넘김')

        for i in range(0, count):
            s = result[i]
            self.req_Real_SC(s['CODE'])
            print('REQ SC : ', s['CODE'])

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
        DATA['DCY'] = self.IndiTR.dynamicCall("GetSingleData(int)", 4)  # 전일대비구분
        DATA['CY'] = self.IndiTR.dynamicCall("GetSingleData(int)", 5)  # 전일대비
        DATA['RCY'] = self.IndiTR.dynamicCall("GetSingleData(int)", 6)  # 전일대비율
        DATA['Vol'] = self.IndiTR.dynamicCall("GetSingleData(int)", 7)  # 누적거래량
        DATA['TRADING_VALUE'] = self.IndiTR.dynamicCall("GetSingleData(int)", 8)  # 누적거래대금
        DATA['ContQty'] = self.IndiTR.dynamicCall("GetSingleData(int)", 9)  # 단위채결량
        DATA['Open'] = self.IndiTR.dynamicCall("GetSingleData(int)", 10)  # 시가
        DATA['High'] = self.IndiTR.dynamicCall("GetSingleData(int)", 11)  # 고가
        DATA['Low'] = self.IndiTR.dynamicCall("GetSingleData(int)", 12)  # 저가
        DATA['POT'] = self.IndiTR.dynamicCall("GetSingleData(int)", 22)  # 거래강도
        DATA['COP'] = self.IndiTR.dynamicCall("GetSingleData(int)", 24)  # 체결강도
        DATA['DOC'] = self.IndiTR.dynamicCall("GetSingleData(int)", 25)  # 체결 매도매수 구분

    def SH(self, stockTime):
        DATA = {}
        DATA['ISIN_CODE'] = self.IndiTR.dynamicCall("GetSingleData(int)", 0)  # 표준코드
        DATA['CODE'] = self.IndiTR.dynamicCall("GetSingleData(int)", 1)  # 단축코드
        DATA['Time'] = self.IndiTR.dynamicCall("GetSingleData(int)", 2)  # 호가접수시간
        DATA['Type'] = self.IndiTR.dynamicCall("GetSingleData(int)", 3) # 장구분
        DATA['SellHOGA'] = []
        DATA['BuyHOGA'] = []
        for i in range(1, 11):
            sellData = [self.IndiTR.dynamicCall("GetSingleData(int)", i*4 + 0),
                        self.IndiTR.dynamicCall("GetSingleData(int)", i*4 + 2)]
            buyData = [self.IndiTR.dynamicCall("GetSingleData(int)", i*4 + 1),
                       self.IndiTR.dynamicCall("GetSingleData(int)", i*4 + 3)]

            DATA['SellHOGA'].append(sellData)
            DATA['BuyHOGA'].append(buyData)

        self.Analyzer.addHoga(DATA['CODE'], stockTime, DATA)

    def AccountList(self):
        count = self.IndiTR.dynamicCall("GetMultiRowCount()")
        result = []

        for i in range(0, count):
            DATA = {}
            DATA['CODE'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 0)  # 단축코드
            DATA['NAME'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 1)  # 종목명
            result.append(DATA)
            print(DATA['CODE'], DATA['NAME'])

        return result

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

        return result

    def order(self):
        DATA = {}
        DATA['Order_Num'] = self.IndiTR.dynamicCall("GetSingleData(int)", 0)  # 주문번호
        DATA['Num'] = self.IndiTR.dynamicCall("GetSingleData(int)", 2)  # 메시지 구분
        DATA['Msg1'] = self.IndiTR.dynamicCall("GetSingleData(int)", 3)  # 메시지1
        DATA['Msg2'] = self.IndiTR.dynamicCall("GetSingleData(int)", 4)  # 메시지2
        DATA['Msg3'] = self.IndiTR.dynamicCall("GetSingleData(int)", 5)  # 메시지3
        print("매수 및 매도 주문결과 :", DATA)

    def rebalancing(self):
        result = self.AccountLookUp()
        sum = 0
        cash_ratio = 0.0

        for i in range(0, len(result)):
            sum += result[i]['CURRENT_PRC'] * result[i]['NUM']

        average = sum / len(result)
        average *= (1 - cash_ratio)

        for i in range(0, len(result)):
            total = result[i]['CURRENT_PRC'] * result[i]['NUM']
            if total > average:
                sell_amount = (total - average) / result[i]['CURRENT_PRC']
                print("판매 : ", result[i]['NAME'], " 개수 : ", sell_amount)
                # TODO : 판매 코드 구현 적용
                # 매수 1호가에 지정가로 던짐
                # self.req_order(result[i]['ISIN_CODE'], result[i]['CURRENT_PRC'], sell_amount, 1, 2)

        for i in range(0, len(result)):
            total = result[i]['CURRENT_PRC'] * result[i]['NUM']
            if total < average:
                buy_amount = (average - total) / result[i]['CURRENT_PRC']
                print("구매 : ", result[i]['NAME'], " 개수 : ", buy_amount)
                # TODO : 판매 코드 구현 적용
                # 매도 1호가에 지정가로 던짐
                # self.req_order(result[i]['ISIN_CODE'], result[i]['CURRENT_PRC'], buy_amount, 2, 2)

    # ==================================================================================

    # TODO : GUI 구현 영역


if __name__ == "__main__":
    app = QApplication(sys.argv)
    Indi = Main()
    Indi.req_stock_mst()
    app.exec_()

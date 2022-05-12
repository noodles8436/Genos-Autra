import os

from PyQt5.QAxContainer import *
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMainWindow
import h5py
import numpy as np
import FileManager

import time as T
import sys


class DataSetMaker(QMainWindow):

    def __init__(self, DATE, path_target, path_result):
        super().__init__()

        self.IndiTR = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")

        self.IndiTR.ReceiveData.connect(self.ReceiveData)
        self.IndiTR.ReceiveSysMsg.connect(self.ReceiveSysMsg)

        self.rqidList = {}

        self.DATE = DATE

        self.read_file = FileManager.configManager(path_target)
        self.write_file = FileManager.configManager(path_result)

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

        self.candleTerm = 3  # 몇분?
        self.candleCount = 20

    def ReceiveData(self, rqid):
        stockCode = self.rqidList[rqid][0]
        time = self.rqidList[rqid][1]
        self.rqidList.__delitem__(rqid)
        self.TR_SCAHRT(stockCode, time)

    def ReceiveSysMsg(self, MsgID):
        print('System Message Received = ', MsgID)
        print('Get Error Message : ', self.IndiTR.GetErrorMessage())

    def extractCandlesByRange(self):

        keys = self.read_file.config.keys()

        print('\n\nTotal SPARK Size : ', len(keys), '\n\n')

        for stockCode in keys:
            stockDatas = self.read_file.config[stockCode]  # List
            print('Found Stock:', stockCode)
            for stockData in stockDatas:
                time = stockData['Time']

                if int(time[:4]) > 1510:
                    continue

                stockSigniture = stockCode + "_" + time
                self.write_file.config[stockSigniture] = []
                self.write_file.config[stockSigniture].append(stockData)
                self.req_TR_SCHART(stockCode, time)
                T.sleep(0.005)

            self.write_file.saveJSON()
            print(stockCode, '추출 및 저장완료\n\n')

        print('\n\n 모든 SPARK 추출 및 저장 완료 ! \n\n')

    def req_TR_SCHART(self, stock_code, time):
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "TR_SCHART")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, stock_code)  # 단축코드
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 1, "1")  # 1:분봉, D:일봉, W:주봉, M:월봉
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 2, str(self.candleTerm))  # 분봉: 1~30, 일/주/월 : 1
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 3, "00000000")  # 시작일(YYYYMMDD)
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 4, "99999999")  # 종료일(YYYYMMDD)
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 5, "9999")  # 조회갯수(1~9999)
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청

        self.rqidList[rqid] = [stock_code, time]

    def TR_SCAHRT(self, stockCode, time):
        received_num = self.IndiTR.dynamicCall("GetMultiRowCount()")
        save_count = self.candleCount

        PREDICT_CANDLE = {}
        PREDICT_CANDLE_ADDED = False

        for i in range(0, received_num):
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

            if int(DATA['TIME']) >= int(time[:4]) + self.candleTerm or DATA['DATE'] != self.DATE:
                PREDICT_CANDLE = DATA
                continue

            if save_count <= 0:
                break

            stockSigniture = stockCode + "_" + time

            if PREDICT_CANDLE_ADDED is False:
                self.write_file.config[stockSigniture].append(PREDICT_CANDLE)
                PREDICT_CANDLE_ADDED = True

            self.write_file.config[stockSigniture].append(DATA)
            save_count -= 1

    def ExtractedFileCheck(self):

        print('무결성 검사 시작')

        write_file = FileManager.configManager(self.write_file.getPATH())
        keys = write_file.config.keys()

        removeList = []

        for key in keys:
            sparkPack = write_file.config[key]
            if len(sparkPack) != 22:
                removeList.append(key)
                print('결함 발견 : ', key)
                continue

        print('\n\n 결함 제거 시작 \n\n')

        for target in removeList:
            write_file.config.__delitem__(target)
            print('삭제되었음 : ', target)

        write_file.saveJSON()
        print('\n\n총 ', len(removeList), '개의 데이터가 삭제되었음')

        self.write_file = write_file
        print('\nWRITE FILE 객체 교체되었음\n')


def MakeHDF5(stockSaveFilePath):
    datasetNum = 0
    ExtractFilePath = stockSaveFilePath + "\\Extract"

    extract_file_list = os.listdir(ExtractFilePath)

    for file in extract_file_list:
        _full_path = ExtractFilePath + "\\" + file
        configFile = FileManager.configManager(_full_path)
        config = configFile.config
        datasetNum += len(config.keys())

    if datasetNum == 0:
        return

    with h5py.File(stockSaveFilePath + "\\DataSet.hdf5", 'w') as f:
        f.create_dataset('image', (datasetNum, 6, 20, 1), dtype='float32')
        f.create_dataset('label', (datasetNum,), dtype='int')

        image_set = f['image']
        label_set = f['label']

        write_index = 0

        for file in extract_file_list:
            _full_path = ExtractFilePath + "\\" + file
            configFile = FileManager.configManager(_full_path)
            config = configFile.config

            for spark in config.keys():
                print('SPARK NAME : ', spark, 'in', file)
                train, target = convertDataToHDF5Array(config[spark])
                image_set[write_index] = train
                label_set[write_index] = target
                write_index += 1

        unique, counts = np.unique(label_set, return_counts=True)
        uniq_cnt_dict = dict(zip(unique, counts))

        print('class Number : ', len(unique))
        print('class counts : ', uniq_cnt_dict)


def convertDataToHDF5Array(sparkData):
    train_result = []
    target_result = 0

    rowList = ['OPEN', 'HIGH', 'Low', 'Close', 'Vol', 'Trading_Value']

    for row in rowList:
        insert_row = []
        for candleData in sparkData[2:]:
            insert_row.append(float(candleData[row]))

        # Scaling
        max = np.max(insert_row)
        min = np.min(insert_row)

        for i in range(0, len(insert_row)):
            if min != max:
                insert_row[i] = round(2 / (max - min) * (insert_row[i] - (max + min) / 2), 6)
            else:
                insert_row[i] = 0

        # Insert
        train_result.append(insert_row)

    train_result = np.array(train_result)
    train_result = train_result.reshape(6, 20, 1)

    target_result = float(sparkData[1]['HIGH']) / int(sparkData[2]['Close'])
    target_result = target_result - 1
    target_result = int(target_result * 100) + 2

    print('>>', target_result)

    return train_result, target_result


def ExtractSparkDatas():
    str_input = input('INPUT DATE (YYYY-MM-DD): ')
    TARGET = "E:\StockSaveFile\\Original\\" + str_input + ".json"
    SAVE_LOCATION = "E:\StockSaveFile\\Extract\\" + str_input + ".json"
    DATE = str_input.replace('-','')
    DataSet_Maker = DataSetMaker(DATE, TARGET, SAVE_LOCATION)
    DataSet_Maker.extractCandlesByRange()


def checkFile():
    str_input = input('INPUT DATE (YYYY-MM-DD): ')
    TARGET = "E:\StockSaveFile\\Original\\" + str_input + ".json"
    SAVE_LOCATION = "E:\StockSaveFile\\Extract\\" + str_input + ".json"
    DATE = str_input.replace('-','')
    DataSet_Maker = DataSetMaker(DATE, TARGET, SAVE_LOCATION)
    DataSet_Maker.ExtractedFileCheck()


def ExtractHDF5():
    MakeHDF5("E:\StockSaveFile")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ExtractSparkDatas()
    #ExtractHDF5()
    print('\n\nDONE\n\n')
    app.exec_()

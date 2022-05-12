import FileManager


class StockRecorder(FileManager.configManager):

    def __init__(self):
        super().__init__()

    def addSpark(self, stockCode, stockData):
        if not self.isSparkExist(stockCode):
            self.config[stockCode] = []
        self.config[stockCode].append(stockData)
        self.saveJSON()

    def addHoga(self, stockCode, Time, hogaData):
        if self.isSparkExist(stockCode):
            SparkList = self.config[stockCode]
            for i in range(0, len(SparkList)):
                if SparkList[i]['Time'] == Time:
                    spark = SparkList[i]
                    spark['Hoga'] = hogaData
                    SparkList[i] = spark
                    self.saveJSON()
                    return

    def isSparkExist(self, stockCode):
        if len(self.config) == 0:
            return False
        if stockCode not in self.config.keys():
            return False
        return True

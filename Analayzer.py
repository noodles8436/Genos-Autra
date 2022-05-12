import StockRecorder


class Analayzer:

    def __init__(self):
        self.focus_list = []
        self.focus_size = 20
        self.TotalStockList = []

        self.StockRecorder = StockRecorder.StockRecorder()

    def check(self, stockData):
        if int(stockData['ContQty']) < 20000:
            return False

        if int(stockData['Time'][:4]) > 1200:
            return False

        self.addFocusStock(stockData['CODE'])
        self.StockRecorder.addSpark(stockData['CODE'], stockData)
        return True

    def addFocusStock(self, StockCode):
        if StockCode in self.focus_list:
            self.focus_list.remove(StockCode)
        self.focus_list.append(StockCode)

        if len(self.focus_list) >= self.focus_size:
            self.focus_list = self.focus_list[:self.focus_size]

        print(self.focus_list)

    def addHoga(self, stockCode, Time, HogaData):
        self.StockRecorder.addHoga(stockCode, Time, HogaData)

    def setTotalStockList(self, _TotalStockList):
        self.TotalStockList = _TotalStockList
        print(self.TotalStockList)
        print("\n\n\n -> 총 ", len(self.TotalStockList), '개\n\n\n')

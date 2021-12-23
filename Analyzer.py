class Analayzer:

    def __init__(self):
        self.focus_list = []
        self.focus_size = 20
        self.TotalStockList = []

    # 주목할 만한 주식이 맞으면 True / 아니면 False 반환
    def check(self, stockData):
        # TODO : 조건검색식
        self.addFocusStock(stockData['CODE'])

    def addFocusStock(self, StockCode):
        if StockCode in self.focus_list:
            self.focus_list.remove(StockCode)
        self.focus_list.append(StockCode)

        if len(self.focus_list) >= self.focus_size:
            self.focus_list = self.focus_list[:self.focus_size]

    def setTotalStockList(self, _TotalStockList):
        self.TotalStockList = _TotalStockList

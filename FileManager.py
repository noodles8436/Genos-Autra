import json
import os
import datetime


class configManager:

    def __init__(self, path=None):
        # if there is no config.json file
        self.config = dict()
        self.PATH = "E:\StockSaveFile\\Original\\" + str(datetime.datetime.now().date()) + ".json"

        if path is not None:
            self.PATH = path

        if os.path.isfile(self.PATH) is not True:
            # Create config.json file
            fd = open(self.PATH, 'w', encoding='UTF-8')
            fd.write("{}")
            fd.close()

        # read config.json and store config.json data in self.config ( dict variable ) using json.load()
        with open(self.PATH, 'r', encoding="UTF-8") as pf:
            self.config = json.load(pf)
            pf.close()

        return

    def getPATH(self):
        return self.PATH

    def saveJSON(self):
        with open(self.PATH, 'w', encoding='UTF-8') as pf:
            json.dump(self.config, pf, indent='\t', ensure_ascii=False)
        pf.close()

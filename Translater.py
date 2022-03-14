from ast import pattern
import os
import pickle
import openpyxl
from AutoCadObject import AutoCad
from os import walk
from Word import Word
from Excel import Excel
import re
from common import AppConfig, initLogging


CONFIG_FILE = "config.ini"
LOGGER = initLogging("翻译器")
try:
    CONFIG = AppConfig(CONFIG_FILE)
except Exception as e:
    LOGGER.error("初始化错误:配置文件中的"+e.args[0]+"找到不值")
    exit(1)

class Translater :

    def __init__(self,dbpath="",dbName=""):
        self.db = HashDatabase(dbpath)
        #中文
        # self.re = re.compile("[\u4e00-\u9fa5]")

    def trans_cad(self):
        app = AutoCad()
        textList = app.searhAll(["AcDbText","AcDbMText"])
        for text in textList:
            en = self.db.query(text.TextString)
            # result = self.re.search(text)
            # tran = self._db.get(result.group(0))
            # self.re.replace(tran)
            if en is not None:
                text.TextString = en

    def trans_word(self):
        word = Word()

        for table in word.iterTables():
            table.trans(self.db)
        
        LOGGER.info("按表翻译完成")
        ############3.1 DLPD(BS201002).00 实瓶输送系统
        # pattern = re.compile("(\d+?.\d+?)\s+?([a-zA-Z0-9\(\)\.\-_]+?)([\s\S]+?)")
        pattern = re.compile("[\u4e00-\u9fa5]+")
        count = 0
        for sentence in word.iterSentences():
            LOGGER.info(sentence.Text)
            if count > int(CONFIG.count):
                break
            count += 1
            text = sentence.Text
            group = pattern.findall(text)
            if group:
                for g in group:
                   en = self.db.query(g)
                   if en is not None:
                       word.replace(g,en)

            # for v in table.iterValues():
            #     value = v.replace("\r\x07","")
            #     en = self.db.query(value)
            #     if en:
            #         word.replace(value,en)

class HashDatabase :


    def __init__(self,dbDir="./db",db_name="hash.db"):
        self.db = dbDir
        self.name = dbDir+"\\"+db_name
        self.map = dict()
        if os.path.exists(self.name):
            self.load()
        else:
            self.reflash()


    def reflash(self):
        for filename in os.listdir(self.db):
            filename:str = filename.lower()
            if filename.endswith("xls") or filename.endswith("xlsx"):
                filePath = self.db +"\\"+filename;
                self.parse_to_map(filePath)

    def parse_to_map(self,file):
        wb = openpyxl.load_workbook(file,read_only=True)
        ws = wb.worksheets[0]
        for row in ws.iter_rows(min_row = 1,values_only=True):
            key = row[0]
            value = row[1]
            self.map[key] = value

    def load(self):
        f = open(self.name,"rb")
        self.map = pickle.load(f)

    def save(self):
        # if os.path.exists(self.name) :
        #     f = os.open(self.name,os.O_WRONLY)
        # else :
        #     f = os.open(self.name,os.O_CREAT|os.O_WRONLY)
        f = open(self.name,"wb")
        pickle.dump(self.map,f)

    def query(self,key):
        return self.map.get(key)
    
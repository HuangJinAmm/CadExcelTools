
import logging
import os
from configobj import ConfigObj

def initLogging(name):
    # create logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # create console handler and set level to debug
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)

    chf = logging.FileHandler("./error.log")
    chf.setLevel(logging.DEBUG)
    # create formatter
    formatter = logging.Formatter('%(asctime)s - %(name)s [ %(levelname)s ] %(message)s')

    # add formatter to ch
    ch.setFormatter(formatter)
    chf.setFormatter(formatter)
    # add ch to logger
    logger.addHandler(ch)
    logger.addHandler(chf)

    return logger

class AppConfig:

    def __init__(self,config):
        ''''''
         #设置相关
        self.config = config
        if not os.path.exists("config.ini"):
            self.createConfig()
        self._config = ConfigObj(config,encoding='UTF-8')
        self.parse()
    
    def setConfig(self, config):
        self.config = config

    def createConfig(self):
        '''初始化配置文件，启动时检查是否存在配置问题就，没有则生成该文件'''
        con=ConfigObj(self.config,encoding='UTF-8')
        con["line_table_setting"]={}
        con["line_table_setting"]["清单列数"]=8
        con["line_table_setting"]["数据填写位置"]="1,1"
        con["line_table_setting"]["模板"]=None
        con["line_table_setting"]["模板数据位置映射"]="1-A,2-B,3-C,4-D,5-E,6-F,7-G,8-H"
        con["line_table_setting"]["输出文件位置"]=None
        con["line_table_setting"]["文件名输入位置"]=None

        con["translate"]["excel"] = "output.xlsx"
        con["translate"]["db"] = "./"
        con["translate"]["count"] = 200
        con.write() 


    def parse(self):
        self.txtCloumn = int(self._config["line_table_setting"]["清单列数"])
        hp = self._config["line_table_setting"]["数据填写位置"]
        self.headPosition = hp.split(",")
        self.templateFile = self._config["line_table_setting"]["模板"]
        cm = self._config["line_table_setting"]["模板数据位置映射"]
        cm = cm.split(",")
        self.cloumnMap = {}
        for c in cm :
            (k,v) = c.split("-")
            self.cloumnMap[int(k)]=v
        self.exportDir = self._config["line_table_setting"]["输出文件位置"]
        self.filenamePos = self._config["line_table_setting"]["文件名输入位置"]

        self.dbDir = self._config["translate"]["db"]
        self.dbName = self._config["translate"]["db_name"]
        self.count = self._config["translate"]["count"]


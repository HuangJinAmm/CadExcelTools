from fileinput import filename
import os
from tkinter import *
from tkinter import messagebox
from PIL import Image,ImageTk
from AutoCadObject import AutoCad
from DataExporter import TxtBomExporter
from Excel import Excel
from Translater import HashDatabase, Translater
from common import AppConfig
from common import initLogging
from trans_tool import parse_example
import win32gui
import win32com


CONFIG_FILE = "config.ini"
LOGGER = initLogging("app")
try:
    CONFIG = AppConfig(CONFIG_FILE)
except Exception as e:
    LOGGER.error("初始化错误:配置文件中的"+e.args[0]+"找到不值")
    exit(1)

def catch_err_to_msgbox(func):
    def wrapper(*args,**kwargs):
        try:
            return func(*args,**kwargs)
        except Exception as e:
            LOGGER.error(e) 
            exceptinfo = e.excepinfo
            if exceptinfo and len(exceptinfo)==6 :
                einfo:str = exceptinfo[2]
                m = exceptinfo[1] + einfo
            else :
                m = str(e)
            messagebox.showwarning(title="错误",message=m);
    return wrapper

class App :

    def __init__(self,size:str,title:str):
        self.window = Tk()
        self.window.title(title)
        self.window.geometry(size)
        self.window.resizable(FALSE,FALSE)

    def createButton(self,name:str,function,image=None):
        btn = Button(self.window,text=name,command=function,height=2,width=48)
        if image is not None:
            btn.configure(image = read_image(image))
        btn.pack()

    def run(self):
        self.window.mainloop()
    
def read_image(filename,height=48,width=48):
    image = Image.open(filename).convert('RGBA')
    image_resize = image.resize((height,width),Image.ANTIALIAS)
    dealed_im= ImageTk.PhotoImage(image_resize)
    return dealed_im

@catch_err_to_msgbox
def tr_word():
    LOGGER.info("开始翻译word")
    trans = Translater(CONFIG.dbDir,CONFIG.dbName)
    trans.trans_word()
    LOGGER.info("完成翻译word")
    messagebox.showinfo(title="完成",message="翻译完成")
    

@catch_err_to_msgbox
def tr_cad():
    LOGGER.info("开始翻译CAD")
    trans = Translater(CONFIG.dbDir,CONFIG.dbName)
    trans.trans_cad()
    LOGGER.info("完成翻译CAD")
    messagebox.showinfo(title="完成",message="翻译完成")

@catch_err_to_msgbox
def exact_BOM():
    try:
        cadApp = AutoCad()
        messagebox.showinfo(title="继续操作",message="请转到CAD选择需要导出的数据！！！")
        cadHWND = cadApp.getHWND
    except Exception as e:
        m = "没有检测到打开的Autocad"
        messagebox.showwarning(title="警告",message=m);
        return

    currentHWND = win32gui.GetForegroundWindow()

    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(cadHWND)

    exporter = TxtBomExporter(cadApp,CONFIG)
    name = cadApp.getCurrentDocName
    headPos = ( int(x) for x in CONFIG.headPosition)
    rows = exporter.export()
    if bool(CONFIG.cloumnMap):
        mapRows = []
        headRow = {}
        begin = int(CONFIG.headPosition[0])
        tuhao = name[0:name.rindex(".")].split(" ")[0]
        tuname = name[0:name.rindex(".")].split(" ")[1]
        headRowValue = [1,tuhao,tuname,1,"组件","","","",""]
        for k,v in CONFIG.cloumnMap.items():
            headRow[v+str(begin)]=headRowValue[k-1]
        begin += 1
        for row in rows:
            rmap = {}
            for k,v in CONFIG.cloumnMap.items():
                if k == 1:
                    rmap[v+str(begin)]="1."+row[k-1]
                else :
                    rmap[v+str(begin)]=row[k-1]
            mapRows.append(rmap)
            begin = begin + 1
    else:
        mapRows=rows
    
    temp = CONFIG.templateFile
    if not os.path.exists(temp):
        temp = None
    excel = Excel(headPosition=headPos,template=temp)
    xlsx = name[0:name.rindex(".")+1]+"xlsx";
    outPath = CONFIG.exportDir
    if bool(outPath) and not outPath.endswith("/"):
        outPath += "\\"
    if os.path.exists(outPath+xlsx) :
        choose = messagebox.askquestion(title="生成文件",message="文件{}已存在是否覆盖?".format(outPath+xlsx))
        if choose == "no" :
            return;
    excel.writeTofile(mapRows,outPath,xlsx,filenamePost=CONFIG.filenamePos)
    m = "导出成功，已保存到"+outPath+xlsx
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(currentHWND)
    messagebox.showinfo(title="完成",message=m)

@catch_err_to_msgbox
def gen_db():
    choose = messagebox.askquestion(title="生成数据",message="是否从doc文件生成?")
    if choose == "yes":
        messagebox.showinfo(title="生成数据库",message="此过程比较耗时，请耐心等候,完成时会弹框提示")
        parse_example(CONFIG.dbDir)
        hashDb = HashDatabase(CONFIG.dbDir,CONFIG.dbName)
        hashDb.save()
        messagebox.showinfo(title="完成",message="完成!!!")
        return
    choose = messagebox.askquestion(title="生成数据",message="是否从XLSX文件刷新数据?")
    if choose == "yes":
        messagebox.showinfo(title="生成数据库",message="此过程比较耗时，请耐心等候,完成时会弹框提示")
        hashDb = HashDatabase(CONFIG.dbDir,CONFIG.dbName)
        hashDb.save()
        messagebox.showinfo(title="完成",message="完成!!!")
        return


@catch_err_to_msgbox
def config():
    LOGGER.info("打开配置")
    os.startfile(CONFIG_FILE)

if __name__ == "__main__":
    app = App("200x230","测试")
    app.createButton("翻译 WORD",tr_word)
    app.createButton("翻译   CAD",tr_cad)
    app.createButton("导出   清单",exact_BOM)
    app.createButton("生成数据库",gen_db)
    app.createButton("配      置",config)
    app.run()
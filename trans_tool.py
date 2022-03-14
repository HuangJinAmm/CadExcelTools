import os
import re
from Excel import Excel
import win32com
from win32com.client import Dispatch

def parse_example(dir:str):
    match = dict()
    result = list()
    en_result = list()
    # re_ch = re.compile("[\u4e00-\u9fa5]+?")
    temp = []
    root = os.getcwd()
    if dir.startswith('.'):
        dir = root +dir[1:len(dir)]
    for filename in os.listdir(dir):
        if filename.endswith("doc") or filename.endswith("docx"):
            f = filename[0:filename.rindex(".")]
            if f.endswith("-en"):
                f = f.replace("-en","")
                if match.get(f,0) == 0:
                    match[f]=[0,0]
                match[f][1]= filename
            else:
                if match.get(f,0) == 0:
                    match[f]=[0,0]
                match[f][0]=filename
    app = Dispatch("word.Application")
    app.Visible = True
    for ch,en in match.values():
        if ch ==0 or en ==0:
            continue
        word_ch = app.Documents.Open(dir+"\\"+ch)
        word_en = app.Documents.Open(dir+"\\"+en)
        for ch_val,en_val in zip(iter_tables_value(word_ch),iter_tables_value(word_en)):
            # if re_ch.match(ch_val):
            t = (ch_val.replace("\r\x07","").replace(" ",""),en_val.replace("\r\x07",""))
            if not t in result:
                result.append(t)
                print("生成：{: <30s}---{: >30s}".format(t[0],t[1]))

        for vlist in iter_tables_row_values(word_en,[2,3]):
            fmt_vlist = [x.replace("\r\x07","") for x in vlist]
            fmt_vlist[0] = fmt_vlist[0].replace(" ","")
            if not fmt_vlist in en_result:
                en_result.append(fmt_vlist)
        word_ch.Close()
        word_en.Close()

    output = Excel()
    output.writeTofile(result,dir+"\\"+"output.xlsx")
    output.writeTofile(en_result,dir+"\\"+"output_en.xlsx")
    app.Quit()
    print("完成")

def iter_tables_value(docx):
    count = docx.Tables.Count
    tables = docx.Tables
    for i in range(1,count+1):
        table = tables.Item(i)
        rows = table.Rows.Count
        for x in range(1,rows+1):
            yield table.Cell(x,3).Range.Text

def iter_tables_row_values(docx,rlist):
    count = docx.Tables.Count
    tables = docx.Tables
    for i in range(1,count+1):
        table = tables.Item(i)
        rows = table.Rows.Count
        for x in range(1,rows+1):
            returnlist=[]
            for c in rlist:
                returnlist.append(table.Cell(x,c).Range.Text)
            yield returnlist
    


if __name__ == "__main__":
    parse_example(os.getcwd()+"\\parse")
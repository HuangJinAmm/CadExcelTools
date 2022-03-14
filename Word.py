import win32com
from win32com.client import Dispatch

class Word:

    def __init__(self,filename="",visible=True):
        self.app = Dispatch("word.Application")
        if filename:
            self.doc = self.app.Documents.Open(filename)
        else:
            self.doc = self.app.Documents.Item(1)
        self.app.Visible = visible
    
    @property
    def contents(self):
        return self.doc.Content.Text

    @property
    def tableCount(self):
        return self.doc.Tables.Count

    def getTable(self,index):
        return  Table(self.doc.Tables.Item(index))

    def iterTables(self):
        for i in range(1,self.tableCount+1):
            yield self.getTable(i)
            
    def iterSentences(self):
        s = self.doc.Sentences
        count = s.Count
        for i in range(1,count):
            yield s.Item(i)

    def replace(self,oldstr:str,newstr:str):
        self.app.Selection.Find.Execute(
        oldstr, False, False, False, False, False, True, 1, False, newstr, 2)


class Table:

    def __init__(self,word_table):
        self._table = word_table

    def getValue(self,row:int,Column:int)->str:
        return self._table.Cell(row,Column).Range.Text

    def setValue(self,row:int,Column:int,newStr:str):
        self._table.Cell(row,Column).Range.Text = newStr

    def trans(self,db):
        for r in range(1,self.rows+1):
            # 处理表头
            if r == 1:
                for c in range(1,self.columns+1):
                    self.trans_cell(r,c,db)
            else :
                self.trans_cell_by_No(r,db)

    def trans_cell(self,r,c,db):
        v = self.getValue(r,c)
        value = v.replace("\r\x07","").replace(" ","")
        en = db.query(value)
        if en:
            self.setValue(r,c,en)

    def trans_cell_by_No(self,r,db):
        v = self.getValue(r,2)
        value = v.replace("\r\x07","").replace(" ","")
        en = db.query(value)
        if en:
            self.setValue(r,3,en)

    @property
    def rows(self):
        return self._table.Rows.Count

    @property
    def columns(self):
        return self._table.Columns.Count

    def iterValues(self):
        for r in range(1,self.rows):
            for c in range(1,self.columns):
                yield self.getValue(r,c)
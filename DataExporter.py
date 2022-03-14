from AutoCadObject import AutoCad
from common import AppConfig

class TxtBomExporter:

    def __init__(self,cadapp:AutoCad,config:AppConfig):
        self.cad = cadapp;
        self.config = config

    def export(self):
        '''导出数据'''
        blist = self.cad.selectOnScreen()
        txtlist = [] # 文本的插入点
        lineList = [] #直线
        for b in blist:
            if b.ObjectName == "AcDbLine":
                lineList.append(b)
            elif b.ObjectName == "AcDbText" or b.ObjectName == "AcDbMText":
                txtlist.append(b)

        #以竖线为分割符
        xSet = [];

        if len(lineList) == 0: return
        for line in lineList:
            if int(line.StartPoint[0]) == int(line.EndPoint[0]):
                x = int(line.StartPoint[0])
                if x in xSet:
                    continue
                else:
                    xSet.append(x)
                
        xSet.sort()

        #按y排序
        txtlist.sort(key=lambda x:x.InsertionPoint[1])
        z1list = []
        z2list = []
        count = self.config.txtCloumn 

        for i,x in enumerate(xSet):
            if i%count == 0:
                z1list.append(x)
        z2list.extend(xSet)
        group = len(z1list)
        
        rowmap = dict()

        total = len(txtlist)
        start_y = txtlist[0].InsertionPoint[1]
        end_y = txtlist[total-1].InsertionPoint[1]

        percent = round((end_y - start_y)/total*5)

        if percent == 0:percent=1

        for txt in txtlist:
            px = int(txt.InsertionPoint[0])
            py = int(txt.InsertionPoint[1])


            # 设置同一行的精度
            py = round(py/(percent))
            # if percent == 1:
            #     while abs(py)/percent < 200:
            #         percent = percent *2

            # 确定此数据位于 x1 组， x2列
            (x1,x2) = self.findColumn(z1list,z2list,px)

            key = x1
            # 初始化字典里的数组
            if key not in rowmap:
                rowmap[key] = dict()
            if py not in rowmap[key]:
                rowmap[key][py] = [None for x in range(self.config.txtCloumn)]

            if x2==self.config.txtCloumn or rowmap[key][py][x2-1] == None:
                rowmap[key][py][x2-1] = txt.TextString 
            # else :
                # rowmap[key][py][x2] = txt.TextString

        rows = []
        for i in range(len(z1list),0,-1):
            column_group = rowmap.get(i)
            for r in column_group.keys():
                rows.append(column_group.get(r))

        return rows

    def findXposition(self,px,z2list):
        tmpx=-1
        col = self.config.txtCloumn
        for i,z2 in enumerate(z2list):
            if px < z2:
                tmpx = i
                break
        return (int(tmpx/col),tmpx%col)

    def findColumn(self,z1list,z2list,px):
        rj = -1
        for j,z1 in enumerate(z1list):
            if px < z1:
                rj = j   
                break
        
        if rj == 0:
            return (0,0)#说明在表格之外

        if rj == -1 :
            rj = len(z1list) # 最右侧一组

        count = (rj-1)*8
        ri = count 
        for i in range(ri,len(z2list)-1):
            if px < z2list[i]:
                ri = i
                break
        
        if ri == count :
            ri = len(z2list) # 最右侧一列

        return (rj,ri%self.config.txtCloumn)

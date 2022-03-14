from msilib.schema import Property
import win32com
from win32com.client import Dispatch,VARIANT
import pythoncom
import re
import datetime
import array
import time
import cad
import operator
import math
import collections

Attribute = collections.namedtuple("Attribute",["index","value"])

#===========================================================
#AutoCad对象构造
#===========================================================
class AutoCad(object):

    def __init__(self):
        self.app = win32com.client.Dispatch('Autocad.Application')
        self.doc = self.app.ActiveDocument
        self.docs = self.app.Documents
        self.layout = self.doc.ActiveLayout
        self.modpace = self.doc.ModelSpace
        self.aselect = self.doc.ActiveSelectionSet

    @property
    def getCurrentDocName(self) -> str:
        return self.doc.Name
        
    @property
    def getHWND(self):
        return self.doc.HWND
    
    def sendCommand(self,cmd):
        self.doc.sendCommand(cmd)

    def postCommand(self,cmd):
        self.doc.PostCommand(cmd)
    # 手动选择
    # @param objecName 对象类型
    # @param nameList 对象名 or关系
    def selectOnScreen(self,objectName="AcDbBlockReference",nameList=[]):
        ft = [100]
        fd = [objectName]
        if bool(nameList):
            ftn = [-4,-4]
            fdn = ['<OR','OR>']
            ft.insert(0,-4)
            fd.insert(0,'<AND')
            for name in nameList:
                ftn.insert(1,2)
                fdn.insert(1,name)
            ft.extend(ftn)
            ft.append(-4)
            fd.extend(fdn)
            fd.append('AND>')
        vft = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, ft)
        vfd = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, fd)
        self.aselect.Clear()
        self.aselect.SelectOnScreen(vft,vfd)
        count =self.aselect.Count
        for i in range(count):
            item = self.aselect.Item(i)
            yield item

    # 整个页面搜索
    def searhAll(self,name,predict = None,limit = None):
        block = self.doc.ActiveLayout.Block
        object_names = name
        if object_names:
            if isinstance(object_names, str):
                object_names = [object_names]
            object_names = [n.lower() for n in object_names]
        if predict == None:
            predict = bool
        count = block.Count
        for i in range(count):
            item = block.Item(i)  
            if limit and i >= limit:
                return
            if object_names:
                object_name = item.ObjectName.lower()
                if not any(possible_name in object_name for possible_name in object_names):
                    continue
            if predict(item):
                yield item  


    # 全选指定对象
    def selectAll(self,typeofcadobj="AcDbBlockReference",namelist = None,predicate = None):
        block = self.doc.ActiveSelectionSet
        block.Clear()
        ft = [100]
        fd = [typeofcadobj]
        if bool(namelist):
            ftn = [-4,-4]
            fdn = ['<OR','OR>']
            ft.insert(0,-4)
            fd.insert(0,'<AND')
            for n in namelist:
                ftn.insert(1,2)
                fdn.insert(1,n)
            ft.extend(ftn)
            ft.append(-4)
            fd.extend(fdn)
            fd.append('AND>')
        vft = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, ft)
        vfd = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, fd)
        block.Select(4,vft,vfd)
        count = block.count
        re_list = []
        for i in range(count):
            item = block.Item(i)
            re_list.append(item)
        return re_list

    # 按两点确定的框选
    def selectByPonits(self,point1,point2,*,objectName="AcDbBlockReference",names="",predict=None):
        ft = [100]
        fd = [objectName]
        if bool(names):
            if isinstance(names,str):
                nameList = [names]
            ftn = [-4,-4]
            fdn = ['<OR','OR>']
            ft.insert(0,-4)
            fd.insert(0,'<AND')
            for name in nameList:
                ftn.insert(1,2)
                fdn.insert(1,name)
            ft.extend(ftn)
            ft.append(-4)
            fd.extend(fdn)
            fd.append('AND>')
        vft = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, ft)
        vfd = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, fd)
        self.aselect.Clear()
        self.aselect.Select(1,point1.toVariant3,point2.toVariant3,vft,vfd)
        count =self.aselect.Count
        if predict == None:
            predict = bool
        for i in range(count):
            item = self.aselect.Item(i)
            if predict(item):
                yield item
#===========================================================
#DwgHandler对象构造
#===========================================================
class DwgHandler:

    def __init__(self,dwg):
        self.dwg = dwg

    @property
    def contents():
        return self.contents
    
#===========================================================
#Dwg对象构造
#===========================================================
class Dwg(AutoCad):

    def __init__(self,cadApp,dwgFrame):
        self.frame = dwgFrame
        boundingBox = self.frame.GetBoundingBox()[0]
        self.point1 = Point(boundingBox[0])
        self.point2 = Point(boundingBox[1])
        self.app = cadApp

    def getTitle(self,title):
        self.title = BlockAttributes(self.app.selectByPonits(self.point1,self.point2,nameList=title)[0])
    
    def getBoms(self,bomName):
        self.boms = self.app.selectByPonits(self.point1,self.point2,nameList=bomName)

    @property
    def readTitle(self):
        self.getTitle()
        return self.title.getAttributes

    @property
    def readBoms(self):
        self.getBoms()
        retList = []
        for bom in self.boms:
            b = BlockAttributes(bom)
            retList.append(b.getAttributes)
        return retList

#===========================================================
#Attribute对象构造
#===========================================================
class BlockAttributes:

    def __init__(self,block):
        self.block = block
        self.initAttr()

    def initAttr(self):
        self.attrMap = dict()
        for a in self.block.getAttributes():
            key = a.TagString
            self.attrMap[key] = a

    @property
    def insertPoint(self):
        return Point(self.block.InsertionPoint)

    @insertPoint.setter
    def insertPoint(self,p1):
        self.block.InsertionPoint = p1.array

    def setAttribute(self,name,value):
        if(self.attrMap.get(name)):
            self.attrMap[name].TextString=value
            return True
        else:
            return False
        
    def setBatchAttributes(self,attributeMap):
        for key,value in attributeMap:
            self.setAttribute(key,value)

    @property
    def getAttributes(self):
        self.initAttr()
        retMap = {}
        for key,a in self.attrMap.items():
            retMap[key] = a.TextString
        return retMap
#===========================================================
#Point对象构造
#===========================================================
class Point:

    def __init__(self,*argv):
        if len(argv) == 1:
            self._x = float(argv[0])
            self._y = 0
            self._z = 0
        elif len(argv) == 2:
            self._x = float(argv[0])
            self._y = float(argv[1])
            self._z = 0 
        elif len(argv) == 3:
            self._x = float(argv[0])
            self._y = float(argv[1])
            self._z = float(argv[2])
        else:
            raise ValueError

    @property
    def x(self):
        return self._x

    @property
    def y(self):
        return self._y
    
    @property
    def z(self):
        return self._z

    @x.setter
    def x(self,value):
        self._x = value

    @y.setter
    def y(self,value):
        self._y = value
    
    @z.setter
    def z(self,value):
        self._z = value
    #@param precision 精度
    # 大于零表示 小数点左边位数
    # 小于零表示小数点右边位数
    def equals2D(self,other,precision:int=0) -> bool:
        if isinstance(other, (list, tuple)):
            return (round(self.x,precision)==round(other[0],precision) and round(self.y,precision)==round(other[1],precision))
        elif isinstance(other, Point):
            return round(self.x,precision)==round(other.x,precision) and round(self.y,precision)== round(other.y,precision) 
        else:
            raise TypeError("cant compare to other"+other.__class__.__name__)

    @staticmethod
    def roundPad(num,precision:int=0):
        if precision <= 0:
            precision = abs(precision)
            return round(num,precision)
        else:
            precision = pow(10,precision)
            return (num // precision)*precision



    @property
    def array(self):
        return array.array('f',[self._x,self._y,self._z])

    @property
    def toVariant3(self):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(self._x,self._y,self._z))

    @property
    def toVariant2(self):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(self._x,self._y))

    def __add__(self, other):
        return self.__left_op(self, other, operator.add)

    def __sub__(self, other):
        return self.__left_op(self, other, operator.sub)

    def __mul__(self, other):
        return self.__left_op(self, other, operator.mul)

    def __div__(self, other):
        return self.__left_op(self, other, operator.truediv)

    __radd__ = __add__
    __rsub__ = __sub__
    __rmul__ = __mul__
    __rdiv__ = __div__
    __floordiv__ = __div__
    __rfloordiv__ = __div__
    __truediv__ = __div__
    _r_truediv__ = __div__

    def __neg__(self):
        return self.__left_op(self, -1, operator.mul)

    def __left_op(self, p1, p2, op):
        if isinstance(p2, (float, int)):
            return Point(op(p1.x, p2), op(p1.y, p2), op(p1.z, p2))
        return Point(op(p1.x, p2.x), op(p1.y, p2.y), op(p1.z, p2.z))

    def __iadd__(self, p2):
        return self.__iop(p2, operator.add)

    def __isub__(self, p2):
        return self.__iop(p2, operator.sub)

    def __imul__(self, p2):
        return self.__iop(p2, operator.mul)

    def __idiv__(self, p2):
        return self.__iop(p2, operator.truediv)

    def __iop(self, p2, op):
        if isinstance(p2, (float, int)):
            self.x = op(self.x, p2)
            self.y = op(self.y, p2)
            self.z = op(self.z, p2)
        else:
            self.x = op(self.x, p2.x)
            self.y = op(self.y, p2.y)
            self.z = op(self.z, p2.z)
        return self

    def __repr__(self):
        return self.__str__()

    def __str__(self):
        return 'Apoint(%.2f, %.2f, %.2f)' % (self.x,self.y,self.z)

    def __eq__(self, other):
        if isinstance(other, (list, tuple)):
            return self.x==other[0] and self.y==other[1] and self.z ==other[2]
        elif isinstance(other, Point):
            return self.x == other.x and self.y ==other.y and self.z == other.z
        else:
            raise TypeError("cant compare to other"+other.__class__.__name__)



def test_selectAll():
    app = AutoCad()
    for item in app.searhAll(["AcDbText","AcDbMText"]):
        print(item.objectName)


if __name__ == "__main__":
    test_selectAll()

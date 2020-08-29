import win32com
from win32com.client import Dispatch,VARIANT
import pythoncom
import re
import datetime
import array
import time
import cad

AUTHOR = '黄金'
TODAY = datetime.date.isoformat(datetime.date.today())
PAT = ".+?\(G?\d{6}\)"
ATT = ("序号","代号","名称","数量","材料","单重","总重","品牌","备注","编码","单位")

def select_In_Cad(doc,typeofcadobj,namelist = None,predicate = None):
    block = doc.ActiveSelectionSet
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
    block.SelectOnScreen(vft,vfd)
    count = block.count
    re_list = []
    for i in range(count):
        item = block.Item(i)
        re_list.append(item)
    return re_list

def select_allin_cad(doc,typeofcadobj,namelist = None,predicate = None):
    block = doc.ActiveSelectionSet
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
def select_item(cad,name_list,predicate = None):
    doc = cad.ActiveDocument
    block = doc.ActiveSelectionSet
    block.Clear()
    block.SelectOnScreen()
    if predicate == None:
        predicate = bool
    object_names = name_list
    if object_names:
        if isinstance(object_names, str):
            object_names = [object_names]
        object_names = [n.lower() for n in object_names]
    count = block.count
    for i in range(count):
        item = block.Item(i)  # it's faster than `for item in block`
        object_name = item.ObjectName.lower()
        if not any(possible_name in object_name for possible_name in object_names):
            continue
        if predicate(item):
            yield item

def shijue(cad):
    doc = cad.ActiveDocument
    '''
    设置颜色
    '''
    color_set = []
    standard1 =numpy.array([32,7,0])
    standard2 =numpy.array([18,7,0])
    for i in range(7):
        color = cad.GetInterfaceObject("AutoCAD.AcCmColor.20")
        color.ColorIndex = i
        color_set.append(color)
    selectobj = doc.ActiveSelectionSet
    selectobj.SelectOnScreen()
    for x in range(selectobj.count):
        b = selectobj.item(x)
        set_color(b,color_set)

def tz_wz(cad):
    doc = cad.ActiveDocument
    '''
    调整明细表文字大小
    '''
    selectobj = doc.ActiveSelectionSet
    selectobj.SelectOnScreen()
    std = [3,0.5,8]
    std1 = [3,0.5,22]
    std2 = [3,0.5,14]
    # standard1 =numpy.array([32,7,0])
    # standard2 =numpy.array([18,7,0])
    for x in range(selectobj.count):
        b = selectobj.item(x)
        sc_wenzi(b,2,std2)
        sc_wenzi(b,1,std1)
        sc_wenzi(b,8,std)
        
def wbt(cad):
    doc = cad.ActiveDocument
    selectobj = doc.ActiveSelectionSet
    selectobj.SelectOnScreen()
    for x in range(selectobj.count):
        b = selectobj.item(x)
        write_mxb(b,'代号','DTLG02.06.09A-'+str(x+2))

def pd_t(it):
    try:
        num = int(it.Textstring)
        if num > 16 :
            return True
        else:
            return False
    except Exception as e:
        return False

# def draw_function(doc,function):
#     doc.addspline(0,0,0)
#     for x in range(100):
        



def main():
    cad = win32com.client.Dispatch('Autocad.Application')
    select_In_Cad(cad,'AcDbBlockReference',['A3','A4','PC_MXB_BLOCK','PC_TITLE_BLOCK'])
    # doc = cad.ActiveDocument
    # shijue(cad)
    # p = doc.Utility.GetPoint()
    # print(p1)
    # # draw_function(doc,sin)
    '''
    动画
    '''
    # mo = cad.ActiveDocument.ModelSpace

    # for i in range(10):
    #     c = mo.AddCircle((0,0,0),i)
    #     time.sleep(1)
    #     c.Delete()

    '''
    更改Text    
    '''
    # doc = cad.ActiveDocument
    # selectobj = doc.ActiveSelectionSet
    # selectobj.SelectOnScreen()
    # for x in range(selectobj.count):
    #     b = selectobj.item(x)
    #     b.Textstring = str(int(b.TextString)+3)

    '''
    保存文件
    '''
    # doc = cad.ActiveDocument
    # selectobj = doc.ActiveSelectionSet
    # selectobj.SelectOnScreen()
    # c = selectobj.item(0)
    # c_atts = c.GetAttributes()
    # c_atts[15].TextString = AUTHOR
    # c_atts[16].TextString = TODAY
    # docname = c_atts[0].TextString + c_atts[1].TextString+'.dwg'
    # path = 'C:\\Users\\user\\Desktop\\新建文件夹\\完成\\'
    # filename = path + docname
    # print(filename)
    # doc.saveas(filename)
    '''
   调整颜色
    ''' 
    # p = cad.ActiveDocument.utility.getpoint()
    # for obj in cad.iter_objects_fast('text'):
        # print(obj.TextString)

    # print_set(cad)
    # wbt(cad)

    # t = select_item(cad,'Text')
    # for i in t:
    #     i.Textstring = str(int(i.TextString)-2)
    #     print(i.TextString)
       

    '''
    批量打印
    查找块，查找对象
    '''
    # blk = find_bolck(doc,'PC_MXB_BLOCK')
    # items = find_item(blk,'line',predicate = is_vertical)
    # for t in items:
    #     print(t.StartPoint,t.Length)
    
    '''
    填写属性
    '''
    # selectobj = doc.ActiveSelectionSet
    # selectobj.SelectOnScreen()
    # for x in range(selectobj.count):
    #     b = selectobj.item(x)
    #     print(type(b))
    #     print(b.ObjectName)
    #     cb = CADblock(b)
    #     print(b.__dict__)
    #     def append(s):
    #         if s:
    #             return s + '/' + '独立清单'
    #         else:
    #             return '独立清单'
    #     write_mxb(b,'备注','外发工程')
    '''
    数字增加，排序
    '''
    # selectobj = doc.ActiveSelectionSet
    # selectobj.SelectOnScreen()
    # for x in range(selectobj.count):
    #     b = selectobj.item(x)
    #     write_mxb(b,'序号',lambda s : int(s)+3)

def sc_wenzi(b,pos,standard):
    xs = b.XScaleFactor
    atts = b.GetAttributes()
    try:
        numstr = len(atts[pos].TextString)
        att_sc = atts[pos].ScaleFactor
        att_h = atts[pos].Height
        if numstr > round(standard[2]*att_sc/standard[1]):
            atts[pos].ScaleFactor = standard[1] 
            atts[pos].Height = standard[0]*xs*standard[2]/numstr
        # x = numpy.array(att_bb[0])
        # y = numpy.array(att_bb[1])
        # st_bb = standard * xs
        # bb = y-x
        # db = bb - st_bb
        # if db[1]>0:
        #     p = st_bb[1]*0.618/bb[1]
        #     atts[pos].Height = att_h * p
        # else:
        #     p = 1
        # if db[0]>0:
        #     p1 = st_bb[0]*0.95/(bb[0]*p)
        #     scale = att_sc * p1
        #     if scale < 0.5:
        #         scale = 0.5
        #         p12 = 1.6 - att_sc 
        #         atts[pos].Height = atts[pos].Height * p12
        #     atts[pos].ScaleFactor = scale 
    except Exception as e:
        print(e)

def is_vertical(line):
    return line.StartPoint[0] == line.EndPoint[0]

def find_item(block ,name_list,predicate = None):
    if predicate == None:
        predicate = bool
    if not bool(block):
        return None
    else:
        object_names = name_list
        if object_names:
            if isinstance(object_names, str):
                object_names = [object_names]
            object_names = [n.lower() for n in object_names]
        count = block.count
        for i in range(count):
            item = block.Item(i)  # it's faster than `for item in block`
            object_name = item.ObjectName.lower()
            if not any(possible_name in object_name for possible_name in object_names):
                continue
            if predicate(item):
                yield item
        


def find_bolck(doc,name):
    for b in range(doc.Blocks.Count):
        block = doc.Blocks.item(b)
        if block.Name == name:
            return block

def write_block_att(block,tag,string):
    for att in block.GetAttributes():
        if att.TagString == tag:
            att.TextString = string
def read_block_atts(block):
    atts = []
    for att in block.GetAttributes():
        atts.append( att.TextString)
    return atts

def write_mxb(block,tag,string):
    atts = block.GetAttributes()
    i = ATT.index(tag)
    if callable(string):
        atts[i].TextString = string(atts[i].Textstring)
    else:
        atts[i].TextString = string

def set_color(block,color_set):
    atts = block.GetAttributes()
    atts[0].TrueColor = color_set[6]
    atts[2].TrueColor = color_set[4]
    atts[3].TrueColor = color_set[3]
    atts[4].TrueColor = color_set[4]
    # atts[9].TrueColor = color_set[4]
    if atts[1].TextString.startswith('QB'):
        atts[1].TrueColor = color_set[2]
    elif atts[1].TextString.startswith('GB'):
        atts[1].TrueColor = color_set[3]
    elif bool(re.match(PAT,atts[1].Textstring)):
        atts[1].TrueColor = color_set[1]
    elif atts[1].TextString.startswith('D'):
        atts[1].TrueColor = color_set[0]
    else:
        atts[1].TrueColor = color_set[4]


def auto_print(doc):
    pass

def print_set(cad):
    doc = cad.ActiveDocument
    layout = doc.ActiveLayout
    ys_name = layout.CanonicalMediaName
    config_name = layout.ConfigName
    print(ys_name) 
    print(config_name) 



if __name__ == '__main__':
    main()
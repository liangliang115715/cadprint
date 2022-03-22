
from pyautocad import Autocad
import win32com.client
import win32print
import pythoncom


#acad = Autocad(create_if_not_exists = True)
acad = win32com.client.Dispatch("AutoCAD.Application.23")   #这一步差别很大   AutoCAD.Application.23为 ProgID



acaddoc = acad.ActiveDocument
acadmod = acaddoc.ModelSpace
layout = acaddoc.layouts.item('Model')
plot = acaddoc.Plot



_PRINTER = win32print.GetDefaultPrinter()
_HPRINTER = win32print.OpenPrinter(_PRINTER)
#_PrinterStatus = 'Warning'




def PrinterStyleSetting():     
    acaddoc.SetVariable('BACKGROUNDPLOT', 0) # 前台打印
    layout.ConfigName = 'RICOH MP C2011' # 选择打印机
    layout.StyleSheet = 'monochrome.ctb' # 选择打印样式
    layout.PlotWithLineweights = False # 不打印线宽
    layout.CanonicalMediaName = 'A3' # 图纸大小这里选择A3
    layout.PlotRotation = 1 # 横向打印
    layout.CenterPlot = True # 居中打印
    layout.PlotWithPlotStyles = True # 依照样式打印
    layout.PlotHidden = False # 隐藏图纸空间对象
    print(layout.GetPlotStyleTableNames()[-1])
    layout.PlotType = 4 
'''
    PlotType (enum类型):
    acDisplay: 按显示的内容打印. 
    acExtents: 按当前选定空间范围内的所有内容打印. 
    acLimits: 打印当前空间范围内的所有内容. 
    acView: 打印由 ViewToPlot 属性命名的视图.
    acWindow: 打印由 SetWindowToPlot 方法指定的窗口中的所有内容.  ******
    acLayout: 打印位于指定纸张尺寸边缘的所有内容，原点从 0,0 坐标计算。 
'''    
    
    


DEFAULT_START_POSITION =(3,3)

DRAWING_SIZE = (598,422)
DRAWING_INTEND = 700



class BackPrint(object):

    _instance = None
    def __new__(cls, *args, **kw):
        if cls._instance is None:
            cls._instance = super(BackPrint, cls).__new__(cls)
        return cls._instance
    def __init__(self,PositionX,PositionY):
        self.x = PositionX
        self.y = PositionY
    @staticmethod
    def APoint(x,y):
        """坐标点转化为浮点数"""
        # 需要两个点的坐标
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x,y))
    def run(self,Scale = 1.0):
        #self.PrinterStyleSetting()
        po1 = self.APoint(self.x * Scale - 1, self.y * Scale)
        po2 = self.APoint(self.x * Scale - 1 + DRAWING_SIZE[0], self.y * Scale + DRAWING_SIZE[1]) # 左下点和右上点
        layout.SetWindowToPlot(po1, po2)
        PrinterStyleSetting()
        plot.PlotToDevice()      

  

class PrintTask:
    def __init__(self,maxPrintPositionArray,startPosition=(DEFAULT_START_POSITION[0],DEFAULT_START_POSITION[1])):
    
        self._PrinterStatus = 'Waiting'
        self.maxPrintPositionArray = maxPrintPositionArray # 此处要进行数据验证
        self.printBasePointArray = []
        self.taskPoint = startPosition
        self.PrintingTaskNumber = 0
        
        
        
    def runtask(self,):
        if not self.printBasePointArray:
            self.printBasePointArray = self.generalPrintBasePointArray(self.maxPrintPositionArray)
        
        for position in self.printBasePointArray:
            #printBasePointArray形式 ： [(,),(,),]
            self.taskPoint = position
            current_task = BackPrint(*position)
            current_task.run()
            
            self.PrintingTaskNumber = len(win32print.EnumJobs(_HPRINTER,0,-1,1))
            #print('ing-> ',self.PrintingTaskNumber,'position',position)

            while self.PrintingTaskNumber >= 5:               
                time.sleep(1)
                self.PrintingTaskNumber = len(win32print.EnumJobs(_HPRINTER,0,-1,1))
            time.sleep(1)
            
 
    def ResumeTask(self,):
        pass
    def generalPrintBasePointArray(self,maxPrintPositionArray): 
        printBasePointArray = []
        next_drawing_xORy_intend = DRAWING_INTEND
        
        current_x = int((self.taskPoint[0] - 4)/ DRAWING_INTEND)*DRAWING_INTEND + DEFAULT_START_POSITION[0]
        current_y = int((self.taskPoint[1] - 4)/DRAWING_INTEND)*DRAWING_INTEND + DEFAULT_START_POSITION[1]
        
        
        #print(current_x,current_y)
        
        for position in maxPrintPositionArray:
            while current_x <= position + DEFAULT_START_POSITION[0]:
                printBasePointArray.append((current_x,current_y))
                current_x += next_drawing_xORy_intend
            current_x = DEFAULT_START_POSITION[0]
            current_y += next_drawing_xORy_intend         
        return printBasePointArray #printBasePointArray形式 ： [(,),(,),]
    
    def getTaskNumber(self,):
        TaskNumber = self.PrintingTaskNumber
        try:
            TaskNumber = len(win32print.EnumJobs(_HPRINTER,0,-1,1))
            return TaskNumber
        except Exception as e:
            return TaskNumber


if __name__ == '__main__':
    #task = PrintTask([25094,10395,]) # 地下室LG层以下钢柱
    #task = PrintTask([27895,]) # 地下室LG层钢柱
    task = PrintTask([27895,],(6194,4)) # 地下室LG层钢柱
    task.runtask()


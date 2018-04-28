# coding=utf-8
import wx
import xlrd
import os
import xlwt
from xlutils.copy import copy;
'''
思路：
    1.获取路径下所有文件，注意 本代码没有异常处理，所有该文件夹下只能放excel文件，否则合并失败
    2.新建一个excel文件，用于存储全部数据
    3.逐个打开excel文件，逐行读取数据，再用一个列表来保存每行数据。最后该列表中会存储所有的数据
    4.向excel文件中逐行写入
'''
  
class excel():
    def get_allfile_msg(file_dir):
        for root, dirs, files in os.walk(file_dir):
            '''
            print(root) #当前目录路径  
            print(dirs) #当前路径下所有子目录  
            print(files) #当前路径下所有非目录子文件 
            '''
            return root, dirs, files


'''
将目录的路径加上'/'和文件名，组成文件的路径
'''
    def get_allfile_url(root, files):
        i = 0
        allFile_url = []
        for f in files:
            i += 1
            if i == len(files):
                break
            file_url = root + '/' + f
            allFile_url.append(file_url)
        return allFile_url


    def all_to_one(root, allFile_url, file_name='allExcel.xls', title=None):
        # 首先在该目录下创建一个excel文件,用于存储所有excel文件的数据
        file_name = root + '/' + file_name
        create_excel(file_name, title)

        list_row_data = []
        for f in allFile_url:
            # 打开excel文件
            print '打开%s文件' % f
            excel = xlrd.open_workbook(f)
            # 根据索引获取sheet，这里是获取第一个sheet
            table = excel.sheet_by_index(0)
            print '该文件行数为：%d，列数为：%d' % (table.nrows,table.ncols)


            # 获取excel文件所有的行
            for i in range(table.nrows):
                # 跳过第零行，一般为表头
                if i == 0:
                    i += 1
                    continue
                row = table.row_values(i)  # 获取整行的值，返回列表
                list_row_data.append(row)

        print '总数据量为%d' % len(list_row_data)
        # 写入all文件
        add_row(list_row_data, file_name)


    # 创建文件名为file_name,表头为title的excel文件
    def create_excel(file_name, title):
        print '创建文件%s' % file_name
        a = xlwt.Workbook()
        # 新建一个sheet
        table = a.add_sheet('sheet1', cell_overwrite_ok=True)
        # 写入数据
        for i in range(len(title)):
            table.write(0, i, title[i])
        a.save(file_name)


    # 向文件中添加n行数据
    def add_row(list_row_data, file_name):
        # 打开excel文件
        allExcel1 = xlrd.open_workbook(file_name)
        sheet = allExcel1.sheet_by_index(0)
        # copy一份文件,准备向它添加内容
        allExcel2 = copy(allExcel1)
        sheet2 = allExcel2.get_sheet(0)

        # 写入数据
        i = 1
        for row_data in list_row_data:
            for j in range(len(row_data)):
                sheet2.write(sheet.nrows + i, j, row_data[j])
            i += 1
        # 保存文件，将原文件覆盖
        allExcel2.save(file_name)
        print '合并完成'
    
class ButtonFrame(wx.Frame):  
    def __init__(self):  
        wx.Frame.__init__(self, None, -1, 'Combine',   
                size=(150, 200))  
        panel = wx.Panel(self, -1)
        wx.StaticText(panel, -1, "min row:", (10, 10))  
        min_row = wx.SpinCtrl(panel, -1, "", (30, 30), (80, -1))  
        min_row.SetRange(0,200) 
        min_row.SetValue(0)
        wx.StaticText(panel, -1, "max row:", (10, 60))  
        max_row = wx.SpinCtrl(panel, -1, "", (30, 80), (80, -1))  
        max_row.SetRange(0,200) 
        max_row.SetValue(8) 
        self.button = wx.Button(panel, -1, "Run", pos=(30, 120))  
        self.Bind(wx.EVT_BUTTON, self.OnClick, self.button)  
        self.button.SetDefault()  
  
    def OnClick(self, event):
        min_row=self.min_row.GetValue()
        min_row=self.max_row.GetValue()
        # 设置文件夹路径，
        file_dir = 'C:/test/24'
        # 获取目录的路径,路径下的目录名，路径下的文件名
        root, dirs, files = get_allfile_msg(file_dir)
        # 拼凑目录路径+文件名,组成文件的路径,列表
        allFile_url = get_allfile_url(root, files)
        # 设置文件名，用于保存数据
        file_name = 'test.xls'
        # 设置excle文件表头
        title = ['A','B','C']
        all_to_one(root, allFile_url, file_name=file_name, title=title)
   
        self.button.SetLabel("Finished")  
          
if __name__ == '__main__':  
 
    app = wx.App()  
    frame = ButtonFrame()  
    frame.Show()  
    app.MainLoop() 

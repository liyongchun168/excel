#encoding=utf-8
import wx
class MyFrame(wx.Frame):
    def __init__(self,parent):
        wx.Frame.__init__(self,parent,id=-1,title='One',size=(300,200))
        panel=wx.Panel(self)
        button1=wx.Button(parent=panel,id=-1,label=u'登录',pos=(30,150))
        button2=wx.Button(parent=panel,id=-1,label=u'取消',pos=(180,150))
        wx.StaticText(panel,-1,u"用户名",(30,35))
        wx.StaticText(panel,-1,u"密 码",(30,95))
        self.Username=wx.TextCtrl(panel,-1,u"请输入用户名",(100,30),(175,-1))
        self.Password=wx.TextCtrl(panel,-1,"...",(100,90),(175,-1),wx.TE_PASSWORD)
        self.Bind(wx.EVT_BUTTON,self.CloseMe,button1)
    def CloseMe(self,evt):
        UserName=self.Username.GetValue()
        PassWord=self.Password.GetValue()
        if (UserName == 'demo') and (PassWord =='demo'):
            self.Destroy()
            frame2=TwoFrame()
            frame2.Show(True)
class TwoFrame(wx.Frame):
    def __init__(self,parent=None):
        wx.Frame.__init__(self,parent,id=-1,title='TWO',size=(500,500))
class MyApp(wx.App):
    def OnInit(self):
        self.frame=MyFrame(parent=None)
        self.SetTopWindow(self.frame)
        self.frame.Show(True)
        return True
if __name__ == '__main__':
    app=MyApp()
    app.MainLoop()

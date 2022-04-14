import wx
from BACKUP import speedup

class MyFrame(wx.Frame):
    def SpeedUpMode(self):
        speedup.SpeedupCode()

    def __init__(self):
        super().__init__(parent=None, title='Bank data Generator')
        #panel
        panel=wx.Panel(self)

        #checkboxes
        self.cb1 = wx.CheckBox(panel, label='Consider Reference Number', pos=(10, 10))
        self.cb2 = wx.CheckBox(panel, label='Create Backup of previous file', pos=(10, 40))
        self.cb3 = wx.CheckBox(panel, label='Show Overwrite and Same File Error', pos=(10, 70))

        # self.Bind(wx.EVT_CHECKBOX, self.onChecked)

        #buttons
        my_btn = wx.Button(panel, label='Back Up everything and SpeedUP', pos=(15, 100))
        my_btn.Bind(wx.EVT_BUTTON,self.SpeedUpMode)
        my_btn2 = wx.Button(panel, label='Save', pos=(15, 130))
        my_btn3 = wx.Button(panel, label='Save and Run', pos=(15, 160))


        self.Centre()
        self.Show()


if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()

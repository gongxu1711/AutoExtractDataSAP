import win32com.client
import sys
import subprocess
import time
from tkinter import*
from tkinter import messagebox
class SapGui():
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(5)

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return
        
        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection("PROD_ERP", True)

        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize

    def sapLogin(self):

        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "100"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "MAILN1"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Abc.0301"
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").SetFocus
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").caretPosition = 2
            self.session.findById("wnd[0]").sendVKey(0)
            try:
                self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").SetFocus()
                self.session.findById("wnd[1]/tbar[0]/btn[0]").Press()
            except:
                pass
            self.session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").SelectedNode = "F00005"
            # self.session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").TopNode = "F00003"
            self.session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").DoubleClickNode("F00005")
            self.session.findById("wnd[0]/usr/radRB5").SetFocus()
            self.session.findById("wnd[0]/usr/radRB5").Select()
            self.session.findById("wnd[0]/usr/cmbP_PMSTY").SetFocus()
            self.session.findById("wnd[0]/usr/cmbP_PMSTY").Key = "3"
            self.session.findById("wnd[0]/usr/ctxtS_PDATE-LOW").Text = ""
            self.session.findById("wnd[0]/usr/ctxtP_EKGRP").Text = "203"
            self.session.findById("wnd[0]/usr/ctxtP_EKGRP").SetFocus()
            self.session.findById("wnd[0]/usr/ctxtP_EKGRP").CaretPosition = 3
            self.session.findById("wnd[0]/tbar[1]/btn[8]").Press()

            self.session.findById("wnd[0]/usr/cntlCONTAINER_1100/shellcont/shell").PressToolbarContextButton("&MB_EXPORT")
            self.session.findById("wnd[0]/usr/cntlCONTAINER_1100/shellcont/shell").SelectContextMenuItem("&XXL")

            self.session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\\Users\\USER\\OneDrive - MB Ageas Life\\MBAL_Dataset"
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "MBAL_Dataset_Expenses.xlsx"
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 26
            self.session.findById("wnd[1]/tbar[0]/btn[11]").Press()
        except:
            print(sys.exc_info()[0])
        messagebox.showinfo('showinfo', 'login successfully')

if __name__ == '__main__':
    window = Tk()
    window.title('AutoTransferFile')
    window.geometry('300x100')
    btn = Button(window, text="Extract Data Expenses", command= lambda: SapGui().sapLogin())
    btn.pack()
    mainloop()

import win32com.client
import sys
import subprocess 
import time

import win32clipboard as clip
import pandas as pd

class SapGui():
    def __init__(self, line, start, end):
        
        self.line = line
        self.start = start
        self.end = end
        
    def open_sap(self):
        self.path = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
        subprocess.Popen(self.path)
        
    def connect_sap(self):
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
        self.connection = app.Children(0)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById('wnd[0]').maximize
        
    def sap_login(self):
        self.open_sap()
        time.sleep(3)
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return
        
        self.app = self.SapGuiAuto.GetScriptingEngine
        self.connection = self.app.OpenConnection("E6P-NA/EU ERP[SSO]", True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize
        
        
        
        try:
            
            self.session.findById("wnd[0]").sendVKey(0)
            
        except:
            print(sys.exc_info()[0])
        print("login successful")
        
    def pull_schedule(self):
        
        clip.OpenClipboard()
        clip.EmptyClipboard()
        clip.CloseClipboard()
        
        self.session.findById("wnd[0]").maximize
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "cooispi"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_EXOP1").selected = True
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/SC LAYOUT " + self.line
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PARBPL-LOW").text = self.line
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_OWERK-LOW").text = "1036"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYOP1").text = "teco"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").text = self.start
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").text = self.end
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").setFocus()
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").caretPosition = 9
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&PC")
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
        self.session.findById("wnd[0]").resizeWorkingPane(267,39,False)
        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
        
        clip.OpenClipboard()
        data = clip.GetClipboardData()
        clip.CloseClipboard()
        
        print("Schedule pulled")
        df = data_to_df(data, 'schedule.txt')
        return df
    
    def pull_confirmations(self):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "cooispi"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOH000"
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/SC LAYOUT " + self.line
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PARBPL-LOW").text = self.line
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_OWERK-LOW").text = "1036"
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-LOW").text = self.start
        self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKST-HIGH").text = self.end
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").setFocus()
        self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").caretPosition = 12
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&PC")
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
        self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        
        clip.OpenClipboard()
        data = clip.GetClipboardData()
        clip.CloseClipboard()
        
        print("Confirmations pulled")
        df = data_to_df(data, 'confirmations.txt')
        return df
        
        
def data_to_df(data_clip, file):
    with open(file, 'w+') as report:
        report.write(data_clip)
    
    df = pd.DataFrame()
    df = pd.read_table(file, sep='|', encoding='UTF-8', skipfooter = 1, engine = 'python',skiprows=range(6))
    df.drop("Unnamed: 0", axis=1, inplace=True)
    df.drop("Unnamed: 15", axis=1, inplace=True)
    df.drop(0, axis=0, inplace=True)
    print("Dataframe complete")
    return df

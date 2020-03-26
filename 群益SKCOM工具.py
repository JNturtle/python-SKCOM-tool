import sys
from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QApplication,QDialog, QMessageBox
from F import *

class SKCOMUpdater(QDialog):
    def __init__(self):
        super(SKCOMUpdater,self).__init__()
        loadUi(r'interface.ui',self)
        self.resetLabel()
        self.Refresh.clicked.connect(self.resetLabel)
        self.Check.clicked.connect(self.Check_clicked)

    def resetLabel(self):
        reset()
        self.OSBit.setText(ID['OSBit'])
        self.lastestVersion.setText(ID['LastestCOM']['version'])
        self.lastestVersionName.setText(ID['LastestCOM']['name'])
        self.lastestVersionPath.setText(ID['LastestCOM']['path'])        
        self.regVersion.setText(ID['regCOM']['version'])
        self.regVersionName.setText(ID['regCOM']['name'])
        self.regVersionPath.setText(ID['regCOM']['path'])       
        self.moduleVersion.setText(ID['CCCOM']['version'])
        self.moduleVersionName.setText(ID['CCCOM']['name'])
        self.moduleVersionPath.setText(ID['CCCOM']['path'])     
        
        pass

    def Check_clicked(self):
        """
        檢查三者版本的差異，並顯示提示視窗
        """
        LCV = ID['LastestCOM']['version']
        RCV = ID['regCOM']['version']
        MCV = ID['CCCOM']['version']
        if LCV != defaultVersion:
            title = "資料夾內的COM元件不是最新版本"
            content = "\n".join(["資料夾內的COM元件 {:}".format(LCV), "不是最新版本 {:}".format(defaultVersion), "請下載最新群益API的COM替換"])
            showMegBox(self, "critical", title, content)
        elif LCV == "不存在":
            title = "找不到資料夾內的COM元件"
            content = "\n".join(["找不到資料夾內的COM元件", "請重新下載本程式"])
            showMegBox(self, "critical", title, content)
        elif RCV == "不存在":
            title = "尚未註冊或找不到COM元件"
            content = "\n".join(["尚未註冊或找不到COM元件", "請使用「系統管理員身分執行」元件中的「install.bat」以註冊COM元件"])
            showMegBox(self, "critical", title, content)
        elif LCV != RCV:
            title = "COM註冊版本不是最新版本"
            content = "\n".join(["COM註冊版本不是最新版本", "請使用「系統管理員身分執行」元件中的「Uninstall.bat」「install.bat」以重新註冊COM元件"])
            showMegBox(self, "critical", title, content)
        elif ID['LastestCOM']['name'] != ID['regCOM']['name']:
            title = "COM註冊版本位元不是正確位元"
            content = "\n".join(["COM註冊版本位元不是正確位元", "請使用「系統管理員身分執行」元件中的「Uninstall.bat」「install.bat」以重新註冊COM元件"])
            showMegBox(self, "critical", title, content)
        elif ID['CCCOM']['name'] != ID['regCOM']['name']:
            title = "COM註冊版本位元與ComtypesClient模型使用的COM不同"
            content = "\n".join(["COM註冊版本位元與ComtypesClient模型使用的COM不同", "請重新註冊COM元件或再次檢查所使用的COM版本"])
            showMegBox(self, "critical", title, content)
            title = "重新生成ComtypesClient模型"
            content = "\n".join(["重新生成ComtypesClient模型？", "若版本有差異，可能會導致群益API無法正常使用！"])
            if showMegBox(self, "info", title, content, YesNo = True): 
                delCCModule()    
                import comtypes.client as cc
                cc.GetModule(ID['LastestCOM']['path'])
        elif RCV != MCV and MCV != "不存在":
            title = "COM註冊版本與ComtypesClient模型版本不同"
            content = "\n".join(["COM註冊版本與ComtypesClient模型版本不同", "請檢查是否程式使用的COM是否為最新版本"])
            showMegBox(self, "critical", title, content)
            title = "刪除舊版本的ComtypesClient模型"
            content = "\n".join(["是否需要刪除舊版本的ComtypesClient模型？", "若版本有差異，可能會導致群益API無法正常使用！"])
            if showMegBox(self, "info", title, content, YesNo = True): delCCModule()            
        # info 等級
        elif MCV == "不存在":
            title = "尚未生成新版COM的py模組"
            content = "\n".join(["尚未生成新版COM模組", "可以正常執行群益API"])
            showMegBox(self, "info", title, content)
            title = "從更新器中產生新版COM的py模組"
            content = "\n".join(["是否從更新器中產生新版COM的py模組？"])
            if showMegBox(self, "info", title, content, YesNo = True): 
                import comtypes.client as cc
                cc.GetModule(ID['LastestCOM']['path'])
        else:
            title = "沒有任何異常"
            content = "\n".join(["沒有檢查到任何錯誤"]) 
            showMegBox(self, "info", title, content)

        self.resetLabel()


def showMegBox(parent, level, title, content, YesNo = False):
    if not YesNo:
        if level == "critical":
            QMessageBox.critical(parent, title, content, QMessageBox.Yes)
        else:
            QMessageBox.information(parent, title, content, QMessageBox.Yes)
    else:
        reply = QMessageBox.information(parent, title, content, QMessageBox.Yes|QMessageBox.No, QMessageBox.No)
        if reply == 16384: return True
        else: return False

if __name__ == '__main__':
    App = QApplication(sys.argv)
    Window = SKCOMUpdater()
    Window.show()
    sys.exit(App.exec_())
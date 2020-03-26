
# 基本功能
import win32api
import win32con
import os
from win32api import GetFileVersionInfo
from win32api import RegQueryValue
from win32con import HKEY_LOCAL_MACHINE
from os import listdir, remove
from os.path import isfile, isdir
from os.path import split as OSPATH_split
from comtypes.client import gen_dir, GetModule
#
defaultVersion = "2.13.20.0"
infoDict = {}
ID = infoDict
#
def getOSBit():
    from platform import architecture
    if architecture()[0] == "64bit": return "x64"
    else: "x84"
def checkVersion(COM_path):
    if not isfile(COM_path): 
        COM_path = "不存在"
        LastestVersion, LastestVersionName = "不存在", "不存在"
    else:
        LastestVersionName = GetFileVersionInfo(COM_path, '\\StringFileInfo\\040904b0\\ProductName')
        LastestVersion = GetFileVersionInfo(COM_path, '\\StringFileInfo\\040904b0\\FileVersion')
    return  COM_path, LastestVersionName, LastestVersion
def getRegCOMPath():
    try:
        return RegQueryValue(HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\WOW6432Node\TypeLib\{75AAD71C-8F4F-4F1F-9AEE-3D41A8C9BA5E}\1.0\0\win32")
    except:
        pass
    return "不存在"   
def getCCModuleCOMPath():
    ModulePyName = "_75AAD71C_8F4F_4F1F_9AEE_3D41A8C9BA5E_0_1_0.py"
    ModulePyPath = "\\".join([gen_dir,ModulePyName])
    if not isfile(ModulePyPath): return "不存在"
    with open(ModulePyPath, 'r') as f:
        f.readline(1000)
        data = f.readline(1000)
        f.close()
    return data[16:-2].replace("\\\\","\\")
def delCCModule():   
    def delFile(path, name):
        for fileName in listdir(path)[:]:
            if fileName.find(name) != -1:             
                filePath = "\\".join([path, fileName])
                print("刪除", filePath)
                remove(filePath)
            elif len(fileName.split(".")) == 1:
                filePath = "\\".join([path, fileName])
                delFile(filePath, name)
    delFile(gen_dir, "_75AAD71C_8F4F_4F1F_9AEE_3D41A8C9BA5E_0_1_0")
def reset():
    infoDict['OSBit'] = getOSBit()
    infoDict['LastestCOM'] = {}
    COM_dir = "\\".join([OSPATH_split(__file__)[0], "元件", infoDict['OSBit']])
    if not isdir(COM_dir): infoDict['LastestCOM']['dir'] ="不存在"
    else: infoDict['LastestCOM']['dir'] = COM_dir

    COM_Names = ['LastestCOM', 'regCOM', 'CCCOM']
    COM_Paths = []
    COM_Paths.append("\\".join([COM_dir, "SKCOM.dll"]))
    COM_Paths.append(getRegCOMPath())
    COM_Paths.append(getCCModuleCOMPath())
    
    for COM_Name, COM_Path in zip(COM_Names, COM_Paths):
        path, name, version = checkVersion(COM_Path)    
        if COM_Name not in infoDict: infoDict[COM_Name] = {}
        infoDict[COM_Name]['path'] = path
        infoDict[COM_Name]['name'] = name
        infoDict[COM_Name]['version'] = version
reset()

# 視窗

from tkinter import Label, Button, Tk
from tkinter import messagebox
def tool_init(parent, font=('標楷體', 10)):
    parent.l0 = Label(parent, text="目前", font=font)
    parent.l1 = Label(parent, text="系統位元:", font=font)
    parent.l2 = Label(parent, text="目前版本:", font=font)
    parent.l3 = Label(parent, text="目前名稱:", font=font)
    parent.l4 = Label(parent, text="目前位置:", font=font)

    parent.l0.place({'x':30, 'y':14})
    parent.l1.place({'x':30, 'y':10+30})
    parent.l2.place({'x':30, 'y':34+30})
    parent.l3.place({'x':30, 'y':58+30})
    parent.l4.place({'x':30, 'y':82+30})

    parent.l1r = Label(parent, text="註冊", font=font)
    parent.l2r = Label(parent, text="目前版本:", font=font)
    parent.l3r = Label(parent, text="目前名稱:", font=font)
    parent.l4r = Label(parent, text="目前位置:", font=font)

    parent.l1r.place({'x':30, 'y':10+142})
    parent.l2r.place({'x':30, 'y':34+142})
    parent.l3r.place({'x':30, 'y':58+142})
    parent.l4r.place({'x':30, 'y':82+142})

    parent.l1m = Label(parent, text="Comtypes Client", font=font)
    parent.l2m = Label(parent, text="目前版本:", font=font)
    parent.l3m = Label(parent, text="目前名稱:", font=font)
    parent.l4m = Label(parent, text="目前位置:", font=font)

    parent.l1m.place({'x':30, 'y':10+254})
    parent.l2m.place({'x':30, 'y':34+254})
    parent.l3m.place({'x':30, 'y':58+254})
    parent.l4m.place({'x':30, 'y':82+254})

    parent.Check = Button(parent, text = "檢查狀況")
    parent.Check["command"] = lambda: Check_clicked(tool)
    parent.Check.place({'x':30+200, 'y':400})
    parent.Refresh = Button(parent, text = "更新標籤")
    parent.Refresh["command"] = lambda: Refresh_clicked(tool)
    parent.Refresh.place({'x':100+200, 'y':400})

    Refresh_clicked(parent)
def Refresh_clicked(parent):
    reset()
    parent.l1['text'] = "系統位元:" + ID['OSBit']
    parent.l2['text'] = "目前版本:" + ID['LastestCOM']['version']
    parent.l3['text'] = "目前名稱:" + ID['LastestCOM']['name']
    parent.l4['text'] = "目前位置:" + ID['LastestCOM']['path']

    parent.l2r['text'] = "目前版本:" + ID['regCOM']['version']
    parent.l3r['text'] = "目前名稱:" + ID['regCOM']['name']
    parent.l4r['text'] = "目前位置:" + ID['regCOM']['path']

    parent.l2m['text'] = "目前版本:" + ID['CCCOM']['version']
    parent.l3m['text'] = "目前名稱:" + ID['CCCOM']['name']
    parent.l4m['text'] = "目前位置:" + ID['CCCOM']['path']  
def Check_clicked(parent):
    """
    檢查三者版本的差異，並顯示提示視窗
    """
    LCV = ID['LastestCOM']['version']
    RCV = ID['regCOM']['version']
    MCV = ID['CCCOM']['version']
    if LCV != defaultVersion:
        title = "資料夾內的COM元件不是最新版本"
        content = "\n".join(["資料夾內的COM元件 {:}".format(LCV), "不是最新版本 {:}".format(defaultVersion), "請下載最新群益API的COM替換"])
        messagebox.showerror(title=title, message=content)
    elif LCV == "不存在":
        title = "找不到資料夾內的COM元件"
        content = "\n".join(["找不到資料夾內的COM元件", "請重新下載本程式"])
        messagebox.showerror(title=title, message=content)
    elif RCV == "不存在":
        title = "尚未註冊或找不到COM元件"
        content = "\n".join(["尚未註冊或找不到COM元件", "請使用「系統管理員身分執行」元件中的「install.bat」以註冊COM元件"])
        messagebox.showerror(title=title, message=content)
    elif LCV != RCV:
        title = "COM註冊版本不是最新版本"
        content = "\n".join(["COM註冊版本不是最新版本", "請使用「系統管理員身分執行」元件中的「Uninstall.bat」「install.bat」以重新註冊COM元件"])
        messagebox.showerror(title=title, message=content)
    elif ID['LastestCOM']['name'] != ID['regCOM']['name']:
        title = "COM註冊版本位元不是正確位元"
        content = "\n".join(["COM註冊版本位元不是正確位元", "請使用「系統管理員身分執行」元件中的「Uninstall.bat」「install.bat」以重新註冊COM元件"])
        messagebox.showerror(title=title, message=content)
    elif ID['CCCOM']['name'] != ID['regCOM']['name'] and ID['CCCOM']['name'] != "不存在" :
        title = "COM註冊版本位元與ComtypesClient模型使用的COM不同"
        content = "\n".join(["COM註冊版本位元與ComtypesClient模型使用的COM不同", "請重新註冊COM元件或再次檢查所使用的COM版本"])
        messagebox.showerror(title=title, message=content)
        title = "重新生成ComtypesClient模型"
        content = "\n".join(["重新生成ComtypesClient模型？", "若版本有差異，可能會導致群益API無法正常使用！"])
        if messagebox.askokcancel(title=title, message=content): 
            delCCModule()    
            GetModule(ID['LastestCOM']['path'])
    elif RCV != MCV and MCV != "不存在":
        title = "COM註冊版本與ComtypesClient模型版本不同"
        content = "\n".join(["COM註冊版本與ComtypesClient模型版本不同", "請檢查是否程式使用的COM是否為最新版本"])
        messagebox.showerror(title=title, message=content)
        title = "刪除舊版本的ComtypesClient模型"        
        content = "\n".join(["是否需要刪除舊版本的ComtypesClient模型？", "若版本有差異，可能會導致群益API無法正常使用！"])
        if messagebox.askokcancel(title=title, message=content): delCCModule()            
    # info 等級
    elif MCV == "不存在":
        title = "尚未生成新版COM的py模組"
        content = "\n".join(["尚未生成新版COM模組", "可以正常執行群益API"])
        messagebox.showinfo(title=title, message=content)
        title = "從更新器中產生新版COM的py模組"
        content = "\n".join(["是否從更新器中產生新版COM的py模組？"])
        if messagebox.askokcancel(title=title, message=content): 
            GetModule(ID['LastestCOM']['path'])
    else:
        title = "沒有任何異常"
        content = "\n".join(["沒有檢查到任何錯誤"]) 
        messagebox.showinfo(title=title, message=content)

    Refresh_clicked(parent)

if __name__ == '__main__':
    tool = Tk()
    tool.title("SKCOM-tool")   
    tool.geometry('640x480')
    tool_init(tool)
    tool.mainloop()
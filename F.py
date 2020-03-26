import win32api
import win32con
import os
#
defaultVersion = "2.13.20.0"
infoDict = {}
ID = infoDict
#
def getOSBit():
    import platform
    if platform.architecture()[0] == "64bit": return "x64"
    else: "x84"
def checkVersion(COM_path):
    if not os.path.isfile(COM_path): 
        COM_path = "不存在"
        LastestVersion, LastestVersionName = "不存在", "不存在"
    else:
        LastestVersionName = win32api.GetFileVersionInfo(COM_path, '\\StringFileInfo\\040904b0\\ProductName')
        LastestVersion = win32api.GetFileVersionInfo(COM_path, '\\StringFileInfo\\040904b0\\FileVersion')
    return  COM_path, LastestVersionName, LastestVersion
def getRegCOMPath():
    try:
        return win32api.RegQueryValue(win32con.HKEY_LOCAL_MACHINE, r"SOFTWARE\Classes\WOW6432Node\TypeLib\{75AAD71C-8F4F-4F1F-9AEE-3D41A8C9BA5E}\1.0\0\win32")
    except:
        pass
    return "不存在"   
def getCCModuleCOMPath():
    import comtypes.client as cc    
    ModulePyName = "_75AAD71C_8F4F_4F1F_9AEE_3D41A8C9BA5E_0_1_0.py"
    ModulePyPath = "\\".join([cc.gen_dir,ModulePyName])
    if not os.path.isfile(ModulePyPath): return "不存在"
    with open(ModulePyPath, 'r') as f:
        f.readline(1000)
        data = f.readline(1000)
        f.close()
    return data[16:-2].replace("\\\\","\\")
def delCCModule():   
    def delFile(path, name):
        for fileName in os.listdir(path)[:]:
            if fileName.find(name) != -1:             
                filePath = "\\".join([path, fileName])
                print("刪除", filePath)
                os.remove(filePath)
            elif len(fileName.split(".")) == 1:
                filePath = "\\".join([path, fileName])
                delFile(filePath, name)
    import comtypes.client as cc   
    delFile(cc.gen_dir, "_75AAD71C_8F4F_4F1F_9AEE_3D41A8C9BA5E_0_1_0")
def reset():
    infoDict['OSBit'] = getOSBit()
    infoDict['LastestCOM'] = {}
    COM_dir = "\\".join([os.path.split(__file__)[0], "元件", infoDict['OSBit']])
    if not os.path.isdir(COM_dir): infoDict['LastestCOM']['dir'] ="不存在"
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
#

if __name__ == "__main__":   
    """查看 infoDict"""
    for key in infoDict:
        print(key)
        if type(infoDict[key]) is dict:
            for item in infoDict[key]:
                print(item, infoDict[key][item])
        else:
            print(infoDict[key])
        print("-"*20)


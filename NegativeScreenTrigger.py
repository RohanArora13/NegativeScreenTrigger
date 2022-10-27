
import psutil
import win32process
import win32gui
import time
#import win32api
import pyautogui


previousWindow = ""
currentWindow = ""

# def getFileDescription(fname):
#     """
#     pass file location
#     Read file description.
#     """
#     propNames = ('Comments', 'InternalName', 'ProductName',
#         'CompanyName', 'LegalCopyright', 'ProductVersion',
#         'FileDescription', 'LegalTrademarks', 'PrivateBuild',
#         'FileVersion', 'OriginalFilename', 'SpecialBuild')

#     props = {'FixedFileInfo': None, 'StringFileInfo': None, 'FileVersion': None}

#     try:
#         # backslash as parm returns dictionary of numeric info corresponding to VS_FIXEDFILEINFO struc
#         fixedInfo = win32api.GetFileVersionInfo(fname, '\\')
#         # props['FixedFileInfo'] = fixedInfo
#         # props['FileVersion'] = "%d.%d.%d.%d" % (fixedInfo['FileVersionMS'] / 65536,
#         #         fixedInfo['FileVersionMS'] % 65536, fixedInfo['FileVersionLS'] / 65536,
#         #         fixedInfo['FileVersionLS'] % 65536)

#         # \VarFileInfo\Translation returns list of available (language, codepage)
#         # pairs that can be used to retreive string info. We are using only the first pair.
#         lang, codepage = win32api.GetFileVersionInfo(fname, '\\VarFileInfo\\Translation')[0]

#         # any other must be of the form \StringfileInfo\%04X%04X\parm_name, middle
#         # two are language/codepage pair returned from above

#         strInfo = {}
        
#         strInfoPath = u'\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, "FileDescription")
#         ## print str_info
#         file_description = win32api.GetFileVersionInfo(fname, strInfoPath)


    # except:
    #     pass

    # return file_description

def active_window_process_name():
    global previousWindow, currentWindow

    pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
    
    # pid[-1] is the most likely to survive last longer
    currentWindow = psutil.Process(pid[-1]).name()

    #print( "file description "+ getFileDescription(psutil.Process(pid[-1]).as_dict().get("exe")) )

    #full_name_current = getFileDescription(psutil.Process(pid[-1]).as_dict().get("exe"))

    if(previousWindow != currentWindow):
        
   
        print(currentWindow)
        time.sleep(0.05)
        if((previousWindow == "terminal.exe" and currentWindow != "terminal.exe") or (currentWindow == "terminal.exe" and previousWindow != "terminal.exe")):
            # trigger the inversion

            # Holds down the alt key
            pyautogui.keyDown("alt")
            pyautogui.keyDown("win")
            pyautogui.press("n")
            pyautogui.keyUp("alt")
            pyautogui.keyUp("win")
            #time.sleep(2)

        
        previousWindow = currentWindow

        # Presses the tab key once


while True:
    # print(GetWindowText(GetForegroundWindow()))
    try:
        active_window_process_name()
    except BaseException as e:
       print(str(e))
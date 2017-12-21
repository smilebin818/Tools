# -*- coding: utf-8-*-

import os, os.path

import win32process, win32event
import win32api



exe_path = 'F:\python\IT5Color_v4.2_FVT-1Re_ECT\SCCopy\Tool'
# exe_file = 'ECT.exe'
exe_file = 'SCCopy_v0.4.1.exe'
ini_file = r'F:\python\IT5Color_v4.2_FVT-1Re_ECT\SCCopy\Own\1_Printer\ExecInfo_Loc.ini'

try :
        # 
        handle = win32process.CreateProcess(os.path.join(exe_path, exe_file),
                None, None, None, 0,
                win32process.CREATE_NO_WINDOW, 
                None , 
                None,
                win32process.STARTUPINFO())

        win32api.ShellExecute(0, 'open', 'F:\python\RunTool.exe', ini_file + ' 0','',1)
        running = True
except Exception, e:
        print "Create Error!"
        handle = None
        running = False

while running :
        rc = win32event.WaitForSingleObject(handle[0], 1000)
        if rc == win32event.WAIT_OBJECT_0:
                running = False
#end while
print "GoodBye"
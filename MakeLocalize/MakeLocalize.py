
# -*- coding: utf-8 -*-
#! python27
# 制作Localize的自动化工具，减少重复的操作（每次都要做一次）
#
import os, re, time, shutil, distutils.dir_util

import win32process, win32event
import win32api

import tkMessageBox

from Tkinter import *

_author_  = "YanBin"
_version_ = "ver1.0"

################################
# 输入一下信息
# · ECT文件名名称
# · ECT文件名称 （例：路径后面的名称[F:\IT5Color_v4.2_FVT-1Re_ECT\C759_11\IT5Color_v4.2.ECT]）
# · 代码路径（Model）
# · 日本给的localize路径
# 读这儿： 第一次执行会报错， 只需要再次执行就OK了
################################
# localize_from_km = ur'F:\Localize_Base\20171206'
# src_model_folder = ur'F:\WinDrv_Src\IT5_Color_v4.2\KMSrc_2.06.73\Driver\Model'
# ECT_folder = 'IT5Color_v4.2_FVT-1Re_ECT'
# ECT_file = 'IT5Color_v4.2.ECT'
ECT_folder = ''
ECT_file = ''
src_model_folder = ''
localize_from_km = ''
SCCopy_folder = ''
Base_ECT = ''

################################
# 读这儿： 第一次执行会报错， 只需要再次执行就OK了
################################

"""
                   _ooOoo_ 
                  o8888888o 
                  88" . "88 
                  (| -_- |) 
                  O\  =  /O 
               ____/`---'\____ 
             .'  \\|     |//  `. 
            /  \\|||  :  |||//  \ 
           /  _||||| -:- |||||-  \ 
           |   | \\\  -  /// |   | 
           | \_|  ''\---/''  |   | 
           \  .-\__  `-`  ___/-. / 
         ___`. .'  /--.--\  `. . __ 
      ."" '<  `.___\_<|>_/___.'  >'"". 
     | | :  `- \`.;`\ _ /`;.`/ - ` : | | 
     \  \ `-.   \_ __\ /__ _/   .-` /  / 
======`-.____`-.___\_____/___.-`____.-'====== 
                   `=---=' 
"""
################################
# 以下代码不可修改
################################
# ECT_folder = os.path.join(os.getcwd(), ECT_folder)
# SCCopy_folder = os.path.join(ECT_folder, 'SCCopy')

own_folder = ''
gen_folder = ''

# 制作新的ECT工程
def make_ect_project():
    global own_folder
    global gen_folder
    FTuple = os.listdir(src_model_folder)
    folder_re = re.compile('.*[FA|PKI]')

    # 获取Own和Gen的model名称
    for f in FTuple:
        f_path = os.path.join(src_model_folder, f)
        if os.path.isdir(f_path) and not folder_re.search(f):
            if '-' in f :
                gen_folder = f
            else:
                own_folder = f

    # 复制基础ECT工程，更改为现在的ECT工程名, 如果已经存在该工程名，就跳过
    if not os.path.exists(ECT_folder):
        distutils.dir_util.copy_tree(Base_ECT, ECT_folder)

    # 更改文件夹名os.path.isdir(os.path.join(root, dirf, df))
    newFName = ''
    for f in os.listdir(ECT_folder):
        if f in ['bin', 'SCCopy', 'Setting'] or not os.path.isdir(os.path.join(ECT_folder, f)):
            continue

        if '-' in f and 'FA' in f:
            newFName = gen_folder + 'FA'
        elif '-' in f and not 'FA' in f:
            newFName = gen_folder
        if not '-' in f and 'FA' in f:
            newFName = own_folder + 'FA'
        elif not '-' in f and not 'FA' in f and not 'PKI' in f:
            newFName = own_folder
        elif 'PKI' in f:
            newFName = own_folder + 'PKI'

        os.rename(os.path.join(ECT_folder, f), os.path.join(ECT_folder, newFName))

        for sf in os.listdir(os.path.join(ECT_folder, newFName)):
            if '.ECT' in sf:
                os.rename(os.path.join(ECT_folder, newFName, sf), os.path.join(ECT_folder, newFName, ECT_file))
                # break
            elif 'INI' == sf:
                shutil.rmtree(os.path.join(ECT_folder, newFName, 'INI'))

    # 删除文件内容（Reference， Target）
    delFolder_re = re.compile('Reference|Target', re.I)
    for root, dirs, files in os.walk(SCCopy_folder):
        for dirf in dirs:
            # print dirf
            if delFolder_re.search(dirf):
                for df in os.listdir(os.path.join(root, dirf)):
                    if os.path.isdir(os.path.join(root, dirf, df)):
                        shutil.rmtree(os.path.join(root, dirf, df))
                # print root + ":" + dirf

# 添加新的localize进入工具
def copy_localize():
    # 日本给的localize放入Target文件夹中
    for f in os.listdir(localize_from_km):
        if 'own' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if r'printer' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, 'Sample_OWN', '1_Printer', 'Target'))
                elif r'fax' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, 'Sample_OWN', '2_Fax', 'Target'))

        if 'gen' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if r'printer' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, 'Sample_GEN', '1_Printer', 'Target'))
                elif r'fax' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, r'Sample_GEN', r'2_Fax', r'Target'))

        if 'pki' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if 'printer' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, 'Sample_PKI', '1_Printer', 'Target'))

    # 工程内部的localize放入Reference文件夹中
    # 把代码中INI所有文件复制到了ECT工程中，但是里面的pcl，ps，xps部分，不是我们想要的，所有我们要找到并删除它们
    for f in os.listdir(src_model_folder):
        f_path = os.path.join(src_model_folder, f)
        if os.path.isdir(f_path):

            fsrc = os.path.join(f_path, 'CUSTOM', 'INI')
            pdl_re = re.compile('pcl|ps|xps', re.I)

            if own_folder == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Sample_OWN', '1_Printer', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Sample_OWN', '1_Printer', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Sample_OWN', '1_Printer', 'Reference', sf))
            elif '%s%s' % (own_folder, 'FA') == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Sample_OWN', '2_Fax', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Sample_OWN', '2_Fax', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Sample_OWN', '2_Fax', 'Reference', sf))
            elif gen_folder == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Sample_GEN', '1_Printer', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Sample_GEN', '1_Printer', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Sample_GEN', '1_Printer', 'Reference', sf))
            elif '%s%s' % (gen_folder, 'FA') == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Sample_GEN', '2_Fax', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Sample_GEN', '2_Fax', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Sample_GEN', '2_Fax', 'Reference', sf))
            elif '%s%s' % (own_folder, 'PKI') == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Sample_PKI', '1_Printer', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Sample_PKI', '1_Printer', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Sample_PKI', '1_Printer', 'Reference', sf))

# 修改文件[ExecInfo***.ini]里面的路径
def modified_execinfo_file():
    filename_re = re.compile('ExecInfo.*\.ini')

    sub_str_re = re.compile('(ExecLog=|ModelInfo=|SCLog=|Target=|Reference=)(.*\\SCCopy)(.+)')

    for folderName, subfolders, filenames in os.walk(SCCopy_folder):
        for filename in filenames:
            if filename_re.search(filename):
                exec_info_file = open(os.path.join(folderName, filename))
                exec_info_file_str = exec_info_file.read()

                exec_info_file = open(os.path.join(folderName, filename), 'w')
                sub_str = sub_str_re.sub(r'\1%s\3' % SCCopy_folder, exec_info_file_str)
                exec_info_file.write(sub_str)
                exec_info_file.close()

def run_sccopy_tool(exe_path, ini_path):
    exe_file = 'SCCopy_v0.4.1.exe'

    try :
        # 启动EXE文件
        handle = win32process.CreateProcess(os.path.join(exe_path, exe_file),
            None, None, None, 0,
            win32process.CREATE_NO_WINDOW, 
            None , 
            None,
            win32process.STARTUPINFO())

        # 执行AU3生成的exe文件
        win32api.ShellExecute(0, 'open', os.path.join(os.getcwd(), 'RunTool.exe'), ini_path,'',1)

        running = True
    except Exception, e:
        print 'SCCopy Create error'
        handle = None
        running = False

    while running :
        rc = win32event.WaitForSingleObject(handle[0], 1000)
        if rc == win32event.WAIT_OBJECT_0:
            running = False
            print ini_path + "-------OK"

def run_ect_tool(ect_path):
    exe_file = 'ECT.exe'

    try :
        # 启动EXE文件
        handle = win32process.CreateProcess(os.path.join(ECT_folder, exe_file),
            None, None, None, 0,
            win32process.CREATE_NO_WINDOW, 
            None , 
            None,
            win32process.STARTUPINFO())

        # 执行AU3编写的EXE
        win32api.ShellExecute(0, 'open', os.path.join(os.getcwd(), 'RunTool.exe'), ect_path,'',1)

        running = True
    except Exception, e:
        print 'ECT Create error'
        handle = None
        running = False

    while running :
        rc = win32event.WaitForSingleObject(handle[0], 1000)
        if rc == win32event.WAIT_OBJECT_0:
            running = False
            print ect_path + "-------OK"

def run_exec_Info_file():

    ini_file = {'own':{'printer':r'Sample_OWN\1_Printer\ExecInfo_Loc.ini', 'fax':r'Sample_OWN\2_Fax\ExecInfo.ini'},
                'gen':{'printer':r'Sample_GEN\1_Printer\ExecInfo_Loc.ini', 'fax':r'Sample_GEN\2_Fax\ExecInfo.ini'},
                 'pki':{'printer':r'Sample_PKI\1_Printer\ExecInfo.ini'}
                }

    target_file = {'own':{'printer':r'Sample_OWN\1_Printer\Target', 'fax':r'Sample_OWN\2_Fax\Target'},
                    'gen':{'printer':r'Sample_GEN\1_Printer\Target', 'fax':r'Sample_GEN\2_Fax\Target'},
                    'pki':{'printer':r'Sample_PKI\1_Printer\Target'}
                    }

    # 根据日本提供的材料，来判断执行哪些ini文件
    for f in os.listdir(localize_from_km):
        if 'own' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if r'printer' in sf.lower():

                    run_sccopy_tool(os.path.join(SCCopy_folder, 'Tool'), os.path.join(SCCopy_folder, ini_file['own']['printer']) + ' 0')

                    distutils.dir_util.copy_tree(os.path.join(SCCopy_folder, target_file['own']['printer']), os.path.join(ECT_folder, own_folder, 'INI'))
                    if os.path.exists(os.path.join(ECT_folder, own_folder, 'INI', '_Log')):
                        shutil.rmtree(os.path.join(ECT_folder, own_folder, 'INI', '_Log'))
                    run_ect_tool(os.path.join(ECT_folder, own_folder, ECT_file) + ' 1')

                elif r'fax' in sf.lower():
                    run_sccopy_tool(os.path.join(SCCopy_folder, 'Tool'), os.path.join(SCCopy_folder, ini_file['own']['fax']) + ' 0')

                    distutils.dir_util.copy_tree(os.path.join(SCCopy_folder, target_file['own']['fax']), os.path.join(ECT_folder, own_folder + 'FA', 'INI'))
                    if os.path.exists(os.path.join(ECT_folder, own_folder + 'FA', 'INI', '_Log')):
                        shutil.rmtree(os.path.join(ECT_folder, own_folder + 'FA', 'INI', '_Log'))
                    run_ect_tool(os.path.join(ECT_folder, own_folder + 'FA', ECT_file) + ' 1')

        if 'gen' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if r'printer' in sf.lower():
                    run_sccopy_tool(os.path.join(SCCopy_folder, 'Tool'), os.path.join(SCCopy_folder, ini_file['gen']['printer']) + ' 0')

                    distutils.dir_util.copy_tree(os.path.join(SCCopy_folder, target_file['gen']['printer']), os.path.join(ECT_folder, gen_folder, 'INI'))
                    if os.path.exists(os.path.join(ECT_folder, gen_folder, 'INI', '_Log')):
                        shutil.rmtree(os.path.join(ECT_folder, gen_folder, 'INI', '_Log'))
                    run_ect_tool(os.path.join(ECT_folder, gen_folder, ECT_file) + ' 1')

                elif r'fax' in sf.lower():
                    run_sccopy_tool(os.path.join(SCCopy_folder, 'Tool'), os.path.join(SCCopy_folder, ini_file['gen']['fax']) + ' 0')

                    distutils.dir_util.copy_tree(os.path.join(SCCopy_folder, target_file['gen']['fax']), os.path.join(ECT_folder, gen_folder + 'FA', 'INI'))
                    if os.path.exists(os.path.join(ECT_folder, gen_folder + 'FA', 'INI', '_Log')):
                        shutil.rmtree(os.path.join(ECT_folder, gen_folder + 'FA', 'INI', '_Log'))
                    run_ect_tool(os.path.join(ECT_folder, gen_folder + 'FA', ECT_file) + ' 1')

        if 'pki' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if 'printer' in sf.lower():
                    run_sccopy_tool(os.path.join(SCCopy_folder, 'Tool'), os.path.join(SCCopy_folder, ini_file['pki']['printer']) + ' 0')

                    distutils.dir_util.copy_tree(os.path.join(SCCopy_folder, target_file['pki']['printer']), os.path.join(ECT_folder, own_folder + 'PKI', 'INI'))
                    if os.path.exists(os.path.join(ECT_folder, own_folder + 'PKI', 'INI', '_Log')):
                        shutil.rmtree(os.path.join(ECT_folder, own_folder + 'PKI', 'INI', '_Log'))
                    run_ect_tool(os.path.join(ECT_folder, own_folder + 'PKI', ECT_file) + ' 1')

def reset():
    e1_entry_var.set('')
    e2_entry_var.set('')
    e3_entry_var.set('')
    e4_entry_var.set('')
    e5_entry_var.set('')

def checkinfo():
    if ECT_folder == '' or ECT_file == '' or src_model_folder == '' or localize_from_km == '' or Base_ECT == '':
        tkMessageBox.showinfo("warning","每项信息都必须正确填写！不能为空")
        button2['state'] = NORMAL
        return 0

    if Base_ECT.split('\\')[-1] == ECT_folder.split('\\')[-1]:
        tkMessageBox.showinfo("warning","新的ECT工程名不能和Base的工程名相同！\n请更改项目[ECT_Project_Name]")
        button2['state'] = NORMAL
        return 0

    if not os.path.exists(Base_ECT):
        tkMessageBox.showinfo("warning","Base_ECT工程的路径不正确（找不到该路径）")
        button2['state'] = NORMAL
        return 0

    runToolIsNotExist = True
    for f in os.listdir(os.getcwd()):
        if f == 'RunTool.exe':
            runToolIsNotExist = False
            break
    if runToolIsNotExist:
        tkMessageBox.showinfo("warning","缺少执行文件[RunTool.exe]\n确保该文件和MakeLocalize.exe处于同目录")
        button2['state'] = NORMAL
        return 0

    # path_re = re.compile(r".*\\\d.*")
    # # print Base_ECT
    # print path_re.search('F:\python\0IT5Color_v4.1_FVT-0_ECT')
    # if path_re.search(src_model_folder) or path_re.search(Base_ECT):
    #     tkMessageBox.showinfo("warning","请确认文件夹名称前面是否有数字")
    #     button2['state'] = NORMAL
    #     return 0

    fa_count = 0
    for f in os.listdir(Base_ECT):
        if os.path.isdir(os.path.join(Base_ECT, f)):
            if 'fa' in f.lower():
                fa_count = fa_count + 1
            if 'fax' in f.lower():
                tkMessageBox.showinfo("warning","Base_ECT里面，请规范命名方式[XXXFA]，和代码里面保存一致风格\n不要存在[XXXFAX]这样的文件夹")
                button2['state'] = NORMAL
                return 0
    if fa_count > 2:
        tkMessageBox.showinfo("warning","Base_ECT工程里面，存在多个FAX文件夹，请删除多余的\n保留Own和Gen各一个对应的FAX")
        button2['state'] = NORMAL
        return 0

    return 1

def getvalueforgui():
    global ECT_folder
    global ECT_file
    global src_model_folder
    global localize_from_km
    global SCCopy_folder
    global Base_ECT

    localize_from_km = e1.get()
    src_model_folder = e2.get()
    ECT_folder = os.path.join(os.getcwd(), e3.get())
    ECT_file = e4.get()
    Base_ECT = e5.get()

    SCCopy_folder = os.path.join(ECT_folder, 'SCCopy')

    button2['state'] = DISABLED

    if not checkinfo():
        return

    # if ECT_folder == '' or ECT_file == '' or src_model_folder == '' or localize_from_km == '':
    #     tkMessageBox.showinfo("warning","每项信息都必须正确填写！不能为空")
    #     button2['state'] = NORMAL
    # else:

    try:
        make_ect_project()
        copy_localize()
        modified_execinfo_file()
        run_exec_Info_file()

        tkMessageBox.showinfo("info","执行完毕")
        # button2['state'] = NORMAL

    except Exception as e:
        tkMessageBox.showinfo("warning","内部有个Bug没有解决，请关闭本程序，再执行就OK了\n（创建新的ECT工程时，第一次执行本程序，总会报错）")

if __name__ == '__main__':
    master = Tk()
    master.title("Localize自动化工具")
    master.geometry('600x180')
    master.resizable(width=False, height=False)

    e1_entry_var = StringVar()
    e2_entry_var = StringVar()
    e3_entry_var = StringVar()
    e4_entry_var = StringVar()
    e5_entry_var = StringVar()

    Label(master, text="Localize_KM_Path：").grid(sticky=E)
    Label(master, text="Src_Model_Path：").grid(sticky=E)
    Label(master, text="ECT_Project_Name：").grid(sticky=E)
    Label(master, text="ECT_File_Name：").grid(sticky=E)
    Label(master, text="ECT_Base_Path：").grid(sticky=E)

    e1_entry_var.set(r'F:\Localize_Base\20171206')
    e2_entry_var.set(r'F:\WinDrv_Src\IT5_Color_v4.2\KMSrc_2.06.73\Driver\Model')
    e3_entry_var.set(r'IT5Color_v4.2_FVT-1Re_ECT')
    e4_entry_var.set(r'IT5Color_v4.2.ECT')
    e5_entry_var.set(r'F:\Base_ECT\IT5Color_v4.1_FVT-0_ECT')

    e1 = Entry(master, width='60', textvariable = e1_entry_var)
    e2 = Entry(master, width='60', textvariable = e2_entry_var)
    e3 = Entry(master, width='60', textvariable = e3_entry_var)
    e4 = Entry(master, width='60', textvariable = e4_entry_var)
    e5 = Entry(master, width='60', textvariable = e5_entry_var)

    e1.grid(row=0, column=1)
    e2.grid(row=1, column=1)
    e3.grid(row=2, column=1)
    e4.grid(row=3, column=1)
    e5.grid(row=4, column=1)

    button1 = Button(master, text='reset', width = 10, command=reset)
    button1.grid(row=5, column=0)

    button2 = Button(master, text='GO', width = 10, command=getvalueforgui)
    button2.grid(row=5, column=1)

    mainloop()

    # make_ect_project()

    # copy_localize()

    # modified_execinfo_file()

    # run_exec_Info_file()

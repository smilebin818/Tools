
# -*- coding: utf-8 -*-
#! python27
# 制作Localize的自动化工具，减少重复的操作（每次都要做一次）
#
import os, re, time, shutil, distutils.dir_util

import win32process, win32event
import win32api

################################
# 输入一下信息
# · ECT文件名名称
# · 代码路径（Model）
# · 日本给的localize路径
################################
ECT_folder = 'IT5Color_v4.2_FVT-1Re_ECT'
ECT_file = 'IT5Color_v4.2.ECT'
src_model_folder = ur'F:\WinDrv_Src\IT5_Color_v4.2\KMSrc_2.06.73\Driver\Model'
localize_from_km = ur'F:\Localize_Base\20171117'

ECT_folder = os.path.join(os.getcwd(), ECT_folder)
SCCopy_folder = os.path.join(ECT_folder, 'SCCopy')

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
        distutils.dir_util.copy_tree(os.path.join(os.getcwd(), 'Base_ECT'), ECT_folder)

    # 更改文件夹名
    ECT_FTuple = os.listdir(ECT_folder)
    for f in ECT_FTuple:
        if 'Own' in f :
            newOwnF_re = re.compile('(Own)(.*)')
            newOwnF = newOwnF_re.sub(r'%s\2' % own_folder, f)
            os.rename(os.path.join(ECT_folder, f), os.path.join(ECT_folder, newOwnF))
            if os.path.exists(os.path.join(ECT_folder, newOwnF, 'INI')):
                shutil.rmtree(os.path.join(ECT_folder, newOwnF, 'INI'))
        elif 'Gen' in f :
            newGenF_re = re.compile('(Gen)(.*)')
            newGenF = newGenF_re.sub(r'%s\2' % gen_folder, f)
            os.rename(os.path.join(ECT_folder, f), os.path.join(ECT_folder, newGenF))
            if os.path.exists(os.path.join(ECT_folder, newGenF, 'INI')):
                shutil.rmtree(os.path.join(ECT_folder, newGenF, 'INI'))
        elif 'PKI' in f:
            os.rename(os.path.join(ECT_folder, f), os.path.join(ECT_folder, own_folder + 'PKI'))
            if os.path.exists(os.path.join(ECT_folder, own_folder + 'PKI', 'INI')):
                shutil.rmtree(os.path.join(ECT_folder, own_folder + 'PKI', 'INI'))

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
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, 'Own', '1_Printer', 'Target'))
                elif r'fax' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, 'Own', '2_Fax', 'Target'))

        if 'gen' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if r'printer' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, 'Gen', '1_Printer', 'Target'))
                elif r'fax' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, r'Gen', r'2_Fax', r'Target'))

        if 'pki' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if 'printer' in sf.lower():
                    distutils.dir_util.copy_tree(os.path.join(localize_from_km, f, sf, r'INI'), os.path.join(SCCopy_folder, 'PKI', '1_Printer', 'Target'))

    # 工程内部的localize放入Reference文件夹中
    # 把代码中INI所有文件复制到了ECT工程中，但是里面的pcl，ps，xps部分，不是我们想要的，所有我们要找到并删除它们
    for f in os.listdir(src_model_folder):
        f_path = os.path.join(src_model_folder, f)
        if os.path.isdir(f_path):

            fsrc = os.path.join(f_path, 'CUSTOM', 'INI')
            pdl_re = re.compile('pcl|ps|xps', re.I)

            if own_folder == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Own', '1_Printer', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Own', '1_Printer', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Own', '1_Printer', 'Reference', sf))
            elif '%s%s' % (own_folder, 'FA') == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Own', '2_Fax', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Own', '2_Fax', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Own', '2_Fax', 'Reference', sf))
            elif gen_folder == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Gen', '1_Printer', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Gen', '1_Printer', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Gen', '1_Printer', 'Reference', sf))
            elif '%s%s' % (gen_folder, 'FA') == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'Gen', '2_Fax', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'Gen', '2_Fax', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'Gen', '2_Fax', 'Reference', sf))
            elif '%s%s' % (own_folder, 'PKI') == f:
                distutils.dir_util.copy_tree(fsrc, os.path.join(SCCopy_folder, 'PKI', '1_Printer', 'Reference'))
                for sf in os.listdir(os.path.join(SCCopy_folder, 'PKI', '1_Printer', 'Reference')):
                    if pdl_re.search(sf):
                        shutil.rmtree(os.path.join(SCCopy_folder, 'PKI', '1_Printer', 'Reference', sf))


# 修改文件[ExecInfo***.ini]里面的路径
def modified_execinfo_file():
    filename_re = re.compile('ExecInfo.*\.ini')
    for folderName, subfolders, filenames in os.walk(SCCopy_folder):
        for filename in filenames:
            if filename_re.search(filename):
                if os.path.join('SCCopy','Own') in folderName:
                    sub_str_re = re.compile('(.*)(path_own)(.*)')
                    exec_info_file = open(os.path.join(folderName, filename))
                    exec_info_file_str = exec_info_file.read()

                    exec_info_file = open(os.path.join(folderName, filename), 'w')
                    sub_str = sub_str_re.sub(r'\1%s\3' % os.path.join(SCCopy_folder,'Own'), exec_info_file_str)
                    exec_info_file.write(sub_str)
                    exec_info_file.close()

                elif os.path.join('SCCopy','Gen') in folderName:
                    sub_str_re = re.compile('(.*)(path_gen)(.*)')
                    exec_info_file = open(os.path.join(folderName, filename))
                    exec_info_file_str = exec_info_file.read()

                    exec_info_file = open(os.path.join(folderName, filename), 'w')
                    sub_str = sub_str_re.sub(r'\1%s\3' % os.path.join(SCCopy_folder,'Gen'), exec_info_file_str)
                    exec_info_file.write(sub_str)
                    exec_info_file.close()

                elif os.path.join('SCCopy','PKI') in folderName:
                    sub_str_re = re.compile('(.*)(path_pki)(.*)')
                    exec_info_file = open(os.path.join(folderName, filename))
                    exec_info_file_str = exec_info_file.read()

                    exec_info_file = open(os.path.join(folderName, filename), 'w')
                    sub_str = sub_str_re.sub(r'\1%s\3' % os.path.join(SCCopy_folder,'PKI'), exec_info_file_str)
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

    ini_file = {'own':{'printer':r'Own\1_Printer\ExecInfo_Loc.ini', 'fax':r'Own\2_Fax\ExecInfo.ini'},
                'gen':{'printer':r'Gen\1_Printer\ExecInfo_Loc.ini', 'fax':r'Gen\2_Fax\ExecInfo.ini'},
                 'pki':{'printer':r'PKI\1_Printer\ExecInfo.ini'}
                }

    target_file = {'own':{'printer':r'Own\1_Printer\Target', 'fax':r'Own\2_Fax\Target'},
                    'gen':{'printer':r'Gen\1_Printer\Target', 'fax':r'Gen\2_Fax\Target'},
                    'pki':{'printer':r'PKI\1_Printer\Target'}
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

                    distutils.dir_util.copy_tree(os.path.join(SCCopy_folder, target_file['own']['fax']), os.path.join(ECT_folder, own_folder + 'FAX', 'INI'))
                    if os.path.exists(os.path.join(ECT_folder, own_folder + 'FAX', 'INI', '_Log')):
                        shutil.rmtree(os.path.join(ECT_folder, own_folder + 'FAX', 'INI', '_Log'))
                    run_ect_tool(os.path.join(ECT_folder, own_folder + 'FAX', ECT_file) + ' 1')

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

                    distutils.dir_util.copy_tree(os.path.join(SCCopy_folder, target_file['gen']['fax']), os.path.join(ECT_folder, gen_folder + 'FAX', 'INI'))
                    if os.path.exists(os.path.join(ECT_folder, gen_folder + 'FAX', 'INI', '_Log')):
                        shutil.rmtree(os.path.join(ECT_folder, gen_folder + 'FAX', 'INI', '_Log'))
                    run_ect_tool(os.path.join(ECT_folder, gen_folder + 'FAX', ECT_file) + ' 1')

        if 'pki' in f.lower():
            for sf in os.listdir(os.path.join(localize_from_km, f)):
                if 'printer' in sf.lower():
                    run_sccopy_tool(os.path.join(SCCopy_folder, 'Tool'), os.path.join(SCCopy_folder, ini_file['pki']['printer']) + ' 0')

                    distutils.dir_util.copy_tree(os.path.join(SCCopy_folder, target_file['pki']['printer']), os.path.join(ECT_folder, own_folder + 'PKI', 'INI'))
                    if os.path.exists(os.path.join(ECT_folder, own_folder + 'PKI', 'INI', '_Log')):
                        shutil.rmtree(os.path.join(ECT_folder, own_folder + 'PKI', 'INI', '_Log'))
                    run_ect_tool(os.path.join(ECT_folder, own_folder + 'PKI', ECT_file) + ' 1')

if __name__ == '__main__':

    make_ect_project()

    copy_localize()

    modified_execinfo_file()

    run_exec_Info_file()

#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
__title__ = '将TargetOpinionMain python项目转换为exe文件'
__author__ = '皮'
__email__ = 'pipisorry@126.com'
"""
from PyInstaller.__main__ import run

if __name__ == '__main__':
    opts = ['MakeLocalize.py', '-F', '-w']
    # opts = ['TargetOpinionMain.py', '-F', '-w']
    # opts = ['TargetOpinionMain.py', '-F', '-w', '--icon=TargetOpinionMain.ico','--upx-dir','upx391w']

    # -F 表示生成单个可执行文件
    # -w 表示去掉控制台窗口，这在GUI界面时非常有用。不过如果是命令行程序的话那就把这个选项删除吧！
    # -p 表示你自己自定义需要加载的类路径，一般情况下用不到
    # -i 表示可执行文件的图标

    run(opts)
    print 'OK'
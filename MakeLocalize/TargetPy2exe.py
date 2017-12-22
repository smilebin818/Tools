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
    run(opts)
    print 'OK'
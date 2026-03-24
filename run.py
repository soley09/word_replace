#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
word_replace 启动入口
可以选择启动 Word校对工具 或 知识库管理器
"""

import os
import sys
import subprocess

def main():
    # 方案二：通过启动参数隐藏 Python 控制台窗口
    startupinfo = None
    if os.name == 'nt':
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE # 隐藏窗口

    subprocess.run([sys.executable, "src/word_reader.py"], startupinfo=startupinfo)

if __name__ == "__main__":
    main()

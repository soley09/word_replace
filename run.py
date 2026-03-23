#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
clawTest 启动入口
可以选择启动 Word校对工具 或 知识库管理器
"""

import os
import sys

def main():
    print("=" * 50)
    print("  clawTest 工具集")
    print("=" * 50)
    print()
    print("请选择要运行的程序：")
    print("  1. Word校对工具（校对错别字）")
    print("  2. 知识库管理器（管理术语和词库）")
    print("  0. 退出")
    print()
    
    choice = input("请输入选项 (1/2/0): ").strip()
    
    if choice == "1":
        print("\n正在启动 Word校对工具...")
        os.system("python src/word_reader.py")
    elif choice == "2":
        print("\n正在启动 知识库管理器...")
        os.system("python src/errorLibrary/readLibrary.py")
    elif choice == "0":
        print("再见！")
    else:
        print("无效选项，请重试")

if __name__ == "__main__":
    main()

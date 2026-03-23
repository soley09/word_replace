#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
知识库更新工具
将 Word_Library.txt 中的新词条同步到 Word_Library.md
"""

import os
import re
from datetime import datetime

def load_txt_entries(txt_path):
    """从txt文件加载词条"""
    entries = {}
    if not os.path.exists(txt_path):
        return entries
    
    with open(txt_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            # 跳过注释和空行
            if not line or line.startswith('#'):
                continue
            # 解析 错误词=正确词
            if '=' in line:
                parts = line.split('=', 1)
                wrong = parts[0].strip()
                right = parts[1].strip()
                if wrong and right:
                    entries[wrong] = right
    return entries

def load_md_entries(md_path):
    """从md文件加载已有词条"""
    entries = {}
    if not os.path.exists(md_path):
        return entries
    
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
        # 匹配 错误词=正确词 格式
        pattern = r'([^\s=]+)=([^\s=]+)'
        matches = re.findall(pattern, content)
        for wrong, right in matches:
            entries[wrong] = right
    return entries

def update_library(txt_path, md_path):
    """更新知识库"""
    print("=" * 50)
    print("知识库更新工具")
    print("=" * 50)
    
    # 加载词条
    txt_entries = load_txt_entries(txt_path)
    md_entries = load_md_entries(md_path)
    
    print(f"txt词条数: {len(txt_entries)}")
    print(f"md词条数: {len(md_entries)}")
    
    # 找出新增词条
    new_entries = {}
    for wrong, right in txt_entries.items():
        if wrong not in md_entries:
            new_entries[wrong] = right
    
    if new_entries:
        print(f"\n新增词条数: {len(new_entries)}")
        print("\n新增内容:")
        for wrong, right in new_entries.items():
            print(f"  {wrong} = {right}")
        
        # 追加到md文件
        with open(md_path, 'a', encoding='utf-8') as f:
            f.write("\n\n### 新增词条\n")
            for wrong, right in new_entries.items():
                f.write(f"{wrong}={right}\n")
        
        print("\n✅ 已更新到 Word_Library.md")
    else:
        print("\n没有新增词条")
    
    print("=" * 50)

if __name__ == "__main__":
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    txt_path = os.path.join(script_dir, "Word_Library.txt")
    md_path = os.path.join(script_dir, "Word_Library.md")
    
    update_library(txt_path, md_path)
    
    print("\n按回车键退出...")
    input()

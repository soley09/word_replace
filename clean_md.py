
import os
import re

def clean_md(file_path):
    if not os.path.exists(file_path):
        return
    
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    new_lines = []
    for line in lines:
        if '=' in line:
            parts = line.split('=')
            if len(parts) >= 2:
                # 移除等号两边的空格，但保留词条内部的空格（如果那是意图的话，但通常不是）
                # 针对用户提到的 0 体=灵体，我们将其合并为 0体=灵体
                wrong = parts[0].strip().replace(' ', '')
                right = parts[1].strip().replace(' ', '')
                new_lines.append(f"{wrong}={right}\n")
            else:
                new_lines.append(line)
        else:
            new_lines.append(line)
            
    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)

if __name__ == "__main__":
    clean_md(r'c:\Users\Administrator\Desktop\word_replace\src\errorLibrary\Word_Library.md')

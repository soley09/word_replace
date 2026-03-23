#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
readLibrary - 知识库阅读器
功能：管理知识库，手工导入词库
"""

import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime

# 尝试导入 python-docx
try:
    import docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

class ReadLibraryApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("知识库管理器 - readLibrary")
        self.root.geometry("1400x700")
        
        # 知识库路径
        # 使用相对路径以提高可移植性
        self.library_dir = os.path.dirname(os.path.abspath(__file__))
        self.proper_library_path = os.path.join(self.library_dir, "Proper_Word_Library.txt")
        self.pending_review_path = os.path.join(self.library_dir, "Pending_Review.txt")
        self.word_library_path = os.path.join(self.library_dir, "Word_Library.txt")
        self.word_library_md_path = os.path.join(self.library_dir, "Word_Library.md")
        
        # 数据
        self.proper_words = []  # 专业术语
        self.import_preview = []  # 导入预览数据
        self.review_items = []  # 待审核候选错词 [(错词, 正词, 是否选中), ...]
        
        # 排序选项
        self.sort_var = tk.StringVar(value="time")  # 术语排序
        self.done_sort_var = tk.StringVar(value="time")  # 词库排序
        
        # 当前tab
        self.current_tab = "terms"
        
        # 加载数据
        self.load_data()
        
        self.create_widgets()
        self.root.mainloop()
    
    def load_data(self):
        """加载知识库数据"""
        # 加载专业术语
        if os.path.exists(self.proper_library_path):
            with open(self.proper_library_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        self.proper_words.append(line)
    
    def save_proper_words(self):
        """保存专业术语"""
        with open(self.proper_library_path, 'w', encoding='utf-8') as f:
            f.write("# 专业术语库\n")
            f.write(f"# 创建时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("# 格式：术语\n\n")
            for word in self.proper_words:
                f.write(f"{word}\n")
    
    def save_word_library(self, entries):
        """保存到正式词库（去重）"""
        # 先读取现有词库
        existing_entries = set()
        if os.path.exists(self.word_library_path):
            with open(self.word_library_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and '=' in line and not line.startswith('#'):
                        parts = line.split('=')
                        if len(parts) == 2:
                            wrong = parts[0].strip()
                            right = parts[1].strip()
                            existing_entries.add((wrong, right))
        
        # 过滤掉已存在的词条
        new_entries = [(w, r) for w, r in entries if (w, r) not in existing_entries]
        
        if not new_entries:
            return 0  # 没有新词条
        
        # 写入新词条
        with open(self.word_library_path, 'a', encoding='utf-8') as f:
            f.write(f"\n# 导入时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            for wrong, right in new_entries:
                f.write(f"{wrong}={right}\n")
        
        return len(new_entries)
    
    # ===== 错词生成映射表（口语/口音相关）=====
    CHAR_MAPPING = {
        # 标相关
        '标': ['彪', '表', '扁', '彬', '波'],
        # 催相关
        '催': ['吹', '崔', '摧'],
        # 识相关
        '识': ['是', '实', '视'],
        # 体相关
        '体': ['休', '本'],
        # 意相关
        '意': ['义', '已'],
        # 潜相关
        '潜': ['前', '浅'],
        # 主相关
        '主': ['祖', '做'],
        # 合相关
        '合': ['和', '活'],
        # 源相关
        '源': ['缘', '元'],
        # 一相关
        '一': ['已', '意'],
        # 灵相关
        '灵': ['零'],
        # 心相关
        '心': ['新', '芯'],
        # 师相关
        '师': ['是', '思'],
        # 观相关
        '观': ['关'],
        # 整相关
        '整': ['正'],
        # 潜相关
        '潜': ['前', '浅'],
        # 集相关
        '集': ['及', '急'],
        # 集相关
        '集': ['及', '急'],
        # 场相关
        '场': ['常', '长'],
        # 纯相关
        '纯': ['春', '存'],
        # 接相关
        '接': ['结', '街'],
        # 连相关
        '连': ['联', '帘'],
        # 肉相关
        '肉': ['如'],
        # 魂相关
        '魂': ['浑', '昏'],
        # 程序相关
        '程': ['成', '乘'],
        # 本源相关
        '本': ['笨'],
        # 河相关
        '河': ['合', '何'],
        # 地相关
        '地': ['的', '得'],
        # 银相关
        '银': ['金', '人'],
        # 星相关
        '星': ['心', '新'],
        # 系相关
        '系': ['戏', '细'],
        # 器相关
        '器': ['气'],
        # 外相关
        '外': ['为'],
        # 间相关
        '间': ['建', '见'],
        # 房相关
        '房': ['防'],
        # 频相关
        '频': ['贫', '品'],
        # 梦相关
        '梦': ['盟', '猛'],
        # 行动相关
        '行': ['形', '性'],
        # 意识相关
        '黏': ['粘'],
        # 觉相关
        '觉': ['决', '绝'],
        # 醒相关
        '醒': ['腥', '兴'],
        # 强相关
        '强': ['墙', '抢'],
        # 执相关
        '执': ['直'],
        # 着相关
        '着': ['这', '哪'],
        # 同相关
        '同': ['铜', '通'],
        # 领相关
        '领': ['另', '零'],
        # 世相关
        '世': ['事', '是'],
        # 幻相关
        '幻': ['换', '缓'],
        # 相相关
        '相': ['想', '向'],
        # 脉相关
        '脉': ['麦', '卖'],
        # 海相关
        '海': ['黑', '还'],
        # 底相关
        '底': ['的'],
        # 攻相关
        '攻': ['工'],
        # 防相关
        '防': ['房'],
        # 宇相关
        '宇': ['雨', '予'],
        # 宙相关
        '宙': ['州', '粥'],
    }
    
    def generate_candidate_errors(self):
        """生成候选错词"""
        self.review_items = []
        
        for word in self.proper_words:
            # 为每个字生成可能的错词
            for i, char in enumerate(word):
                if char in self.CHAR_MAPPING:
                    for wrong_char in self.CHAR_MAPPING[char]:
                        wrong_word = word[:i] + wrong_char + word[i+1:]
                        # 排除与原词相同的
                        if wrong_word != word:
                            # 添加到候选列表（默认不选中）
                            self.review_items.append((wrong_word, word, False))
        
        # 去重（基于错词和正词的组合）
        seen = set()
        unique_items = []
        for item in self.review_items:
            key = (item[0], item[1])
            if key not in seen:
                seen.add(key)
                unique_items.append(item)
        
        self.review_items = unique_items
        
        # 刷新列表
        self.refresh_review_list()
        
        # 显示结果
        count = len(self.review_items)
        messagebox.showinfo("生成完成", f"已生成 {count} 个候选错词\n\n请在列表中勾选要导入的词条")
    
    def generate_from_selected(self):
        """从选中的术语生成候选错词"""
        # 获取选中的术语
        selected_indices = self.term_listbox.curselection()
        
        if not selected_indices:
            messagebox.showwarning("警告", "请先在术语列表中选中要生成错词的术语\n\n（单击选中单个，Ctrl+单击选中多个）")
            return
        
        # 获取原始列表（未排序的）
        original_words = self.proper_words
        
        # 直接使用原始列表的索引
        selected_words = [original_words[i] for i in selected_indices if i < len(original_words)]
        
        if not selected_words:
            return
        
        # 清空现有候选并只生成选中的
        self.review_items = []
        
        # 生成候选错词
        for word in selected_words:
            for i, char in enumerate(word):
                if char in self.CHAR_MAPPING:
                    for wrong_char in self.CHAR_MAPPING[char]:
                        wrong_word = word[:i] + wrong_char + word[i+1:]
                        if wrong_word != word:
                            self.review_items.append((wrong_word, word, False))
        
        # 去重
        seen = set()
        unique_items = []
        for item in self.review_items:
            key = (item[0], item[1])
            if key not in seen:
                seen.add(key)
                unique_items.append(item)
        
        self.review_items = unique_items
        
        # 刷新并切换到待审核标签
        self.refresh_review_list()
        self.switch_tab("review")
        
        count = len(self.review_items)
        messagebox.showinfo("生成完成", f"已为选中的 {len(selected_words)} 个术语生成 {count} 个候选错词")
    
    def refresh_review_list(self):
        """刷新待审核列表"""
        self.review_listbox.delete(0, tk.END)
        for wrong, right, selected in self.review_items:
            display = f"{'☑' if selected else '☐'} {wrong} → {right}"
            self.review_listbox.insert(tk.END, display)
    
    def create_widgets(self):
        """创建界面组件"""
        
        # ===== 第一行按钮：手工导入 =====
        row1_frame = tk.Frame(self.root, pady=5)
        row1_frame.pack(fill=tk.X)
        
        tk.Label(row1_frame, text="词库导入:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=10)
        
        # 手工导入词库按钮
        tk.Button(
            row1_frame, 
            text="📥 手工导入词库", 
            command=self.import_word_library,
            font=("Microsoft YaHei", 10),
            padx=12,
            pady=3,
            bg="#2196F3",
            fg="white"
        ).pack(side=tk.LEFT, padx=5)
        
        # 更新导入词库按钮
        self.update_btn = tk.Button(
            row1_frame, 
            text="✅ 更新导入词库", 
            command=self.update_word_library,
            font=("Microsoft YaHei", 10),
            padx=12,
            pady=3,
            bg="#4CAF50",
            fg="white",
            state=tk.DISABLED
        )
        self.update_btn.pack(side=tk.LEFT, padx=5)
        
        # ===== 第二行按钮：专业术语 + 标签页 =====
        row2_frame = tk.Frame(self.root, pady=5)
        row2_frame.pack(fill=tk.X)
        
        tk.Label(row2_frame, text="专业术语:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=10)
        
        # 添加专业术语按钮
        tk.Button(
            row2_frame,
            text="➕ 添加专业术语",
            command=self.add_proper_word,
            font=("Microsoft YaHei", 10),
            padx=12,
            pady=3
        ).pack(side=tk.LEFT, padx=5)
        
        # 选中术语生成候选错词按钮
        tk.Button(
            row2_frame,
            text="🎯 选中生成错词",
            command=self.generate_from_selected,
            font=("Microsoft YaHei", 9),
            padx=8,
            pady=3,
            bg="#2196F3",
            fg="white"
        ).pack(side=tk.LEFT, padx=5)
        
        # 排序选项
        tk.Label(row2_frame, text="  排序:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=10)
        
        self.sort_var = tk.StringVar(value="time")
        
        tk.Radiobutton(row2_frame, text="添加顺序", variable=self.sort_var, value="time",
                       font=("Microsoft YaHei", 9), command=self.refresh_term_list).pack(side=tk.LEFT, padx=3)
        tk.Radiobutton(row2_frame, text="字母顺序", variable=self.sort_var, value="alpha",
                       font=("Microsoft YaHei", 9), command=self.refresh_term_list).pack(side=tk.LEFT, padx=3)
        
        # 标签页 - 放在 row2_frame 中
        tk.Label(row2_frame, text="  查看:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=20)
        
        self.term_btn = tk.Button(
            row2_frame,
            text="📚 术语",
            command=lambda: self.switch_tab("terms"),
            font=("Microsoft YaHei", 10),
            padx=10,
            width=10
        )
        self.term_btn.pack(side=tk.LEFT, padx=2)
        
        self.pending_btn = tk.Button(
            row2_frame,
            text="📋 预览",
            command=lambda: self.switch_tab("pending"),
            font=("Microsoft YaHei", 10),
            padx=10,
            width=10
        )
        self.pending_btn.pack(side=tk.LEFT, padx=2)
        
        # 待审核知识库标签页
        self.review_btn = tk.Button(
            row2_frame,
            text="⏳ 待审核",
            command=lambda: self.switch_tab("review"),
            font=("Microsoft YaHei", 10),
            padx=10,
            width=12
        )
        self.review_btn.pack(side=tk.LEFT, padx=2)
        
        # 全部生成候选错词按钮
        tk.Button(
            row2_frame,
            text="🎯 全部生成候选错词",
            command=self.generate_candidate_errors,
            font=("Microsoft YaHei", 9),
            padx=8,
            pady=3,
            bg="#FF9800",
            fg="white"
        ).pack(side=tk.LEFT, padx=10)
        
        self.done_btn = tk.Button(
            row2_frame,
            text="✅ 词库",
            command=lambda: self.switch_tab("done"),
            font=("Microsoft YaHei", 10),
            padx=10,
            width=10
        )
        self.done_btn.pack(side=tk.LEFT, padx=2)
        
        # 词库排序选项
        tk.Label(row2_frame, text="  排序:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=10)
        
        tk.Radiobutton(row2_frame, text="添加顺序", variable=self.done_sort_var, value="time",
                       font=("Microsoft YaHei", 9), command=self.refresh_done_list).pack(side=tk.LEFT, padx=3)
        tk.Radiobutton(row2_frame, text="字母顺序", variable=self.done_sort_var, value="alpha",
                       font=("Microsoft YaHei", 9), command=self.refresh_done_list).pack(side=tk.LEFT, padx=3)
        
        self.view_btn = tk.Button(
            row2_frame,
            text="📖 文档",
            command=lambda: self.switch_tab("view"),
            font=("Microsoft YaHei", 10),
            padx=10,
            width=10
        )
        self.view_btn.pack(side=tk.LEFT, padx=2)
        
        # ===== 列表区域 =====
        list_frame = tk.Frame(self.root, padx=10, pady=10)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 术语列表
        self.term_listbox = tk.Listbox(
            list_frame,
            font=("Microsoft YaHei", 11),
            selectmode=tk.EXTENDED
        )
        
        # 导入预览列表
        self.pending_listbox = tk.Listbox(
            list_frame,
            font=("Microsoft YaHei", 11),
            selectmode=tk.EXTENDED
        )
        
        # 已通过列表
        self.done_listbox = tk.Listbox(
            list_frame,
            font=("Microsoft YaHei", 11),
            selectmode=tk.EXTENDED
        )
        
        # 待审核知识库列表
        self.review_listbox = tk.Listbox(
            list_frame,
            font=("Microsoft YaHei", 11),
            selectmode=tk.EXTENDED
        )
        
        # 查看词库
        self.view_text = scrolledtext.ScrolledText(
            list_frame,
            wrap=tk.WORD,
            font=("Microsoft YaHei", 10)
        )
        
        # ===== 底部操作按钮 =====
        action_frame = tk.Frame(self.root, pady=10)
        action_frame.pack(fill=tk.X, padx=10)
        
        tk.Button(
            action_frame,
            text="全选",
            command=self.select_all,
            font=("Microsoft YaHei", 10),
            padx=15
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            action_frame,
            text="取消全选",
            command=self.deselect_all,
            font=("Microsoft YaHei", 10),
            padx=15
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            action_frame,
            text="🗑️ 删除选中",
            command=self.delete_selected,
            font=("Microsoft YaHei", 10),
            padx=15,
            fg="red"
        ).pack(side=tk.LEFT, padx=5)
        
        # 待审核知识库的批量导入按钮
        tk.Button(
            action_frame,
            text="📥 批量导入词库",
            command=self.batch_import_to_library,
            font=("Microsoft YaHei", 10),
            padx=15,
            bg="#4CAF50",
            fg="white"
        ).pack(side=tk.LEFT, padx=20)
        
        # 状态栏
        self.status_label = tk.Label(
            self.root,
            text=f"专业术语: {len(self.proper_words)} 条 | 导入预览: {len(self.import_preview)} 条 | 待审核: 0 条 | 已选: 0 条",
            font=("Microsoft YaHei", 9),
            fg="gray",
            anchor=tk.W
        )
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        # 初始显示
        self.switch_tab("terms")
        
        # 绑定双击事件（术语列表）
        self.term_listbox.bind('<Double-Button-1>', self.on_term_double_click)
        
        # 绑定点击事件（待审核列表：点击切换选择状态）
        self.review_listbox.bind('<Button-1>', self.on_review_item_click)
    
    def switch_tab(self, tab_name):
        """切换标签页"""
        self.current_tab = tab_name
        
        # 更新按钮样式
        bg_normal = "#E0E0E0"
        bg_active = "#BBDEFB"
        
        self.term_btn.config(bg=bg_active if tab_name == "terms" else bg_normal)
        self.pending_btn.config(bg=bg_active if tab_name == "pending" else bg_normal)
        self.review_btn.config(bg=bg_active if tab_name == "review" else bg_normal)
        self.done_btn.config(bg=bg_active if tab_name == "done" else bg_normal)
        self.view_btn.config(bg=bg_active if tab_name == "view" else bg_normal)
        
        # 隐藏所有列表
        self.term_listbox.pack_forget()
        self.pending_listbox.pack_forget()
        self.review_listbox.pack_forget()
        self.done_listbox.pack_forget()
        self.view_text.pack_forget()
        
        # 添加滚动条
        scrollbar = tk.Scrollbar(self.root)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        if tab_name == "terms":
            self.term_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=self.term_listbox.yview)
            self.term_listbox.config(yscrollcommand=scrollbar.set)
            self.refresh_term_list()
        elif tab_name == "pending":
            self.pending_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=self.pending_listbox.yview)
            self.pending_listbox.config(yscrollcommand=scrollbar.set)
            self.refresh_pending_list()
        elif tab_name == "review":
            self.review_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=self.review_listbox.yview)
            self.review_listbox.config(yscrollcommand=scrollbar.set)
            self.refresh_review_list()
        elif tab_name == "done":
            self.done_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=self.done_listbox.yview)
            self.done_listbox.config(yscrollcommand=scrollbar.set)
            self.refresh_done_list()
        elif tab_name == "view":
            self.view_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=self.view_text.yview)
            self.view_text.config(yscrollcommand=scrollbar.set)
            self.refresh_view()
        
        # 更新状态栏
        self.update_status_label()
    
    def on_review_item_click(self, event):
        """点击待审核列表项：切换选择状态"""
        # 获取点击的行号
        index = self.review_listbox.index(f"@{event.x},{event.y}")
        if index is not None:
            idx = int(index)
            if 0 <= idx < len(self.review_items):
                wrong, right, selected = self.review_items[idx]
                # 切换选中状态
                self.review_items[idx] = (wrong, right, not selected)
                self.refresh_review_list()
                self.update_status_label()
    
    def refresh_term_list(self):
        """刷新术语列表"""
        self.term_listbox.delete(0, tk.END)
        
        # 根据排序方式显示
        if self.sort_var.get() == "alpha":
            # 字母顺序排序
            sorted_words = sorted(self.proper_words, key=lambda x: x.lower())
        else:
            # 按添加顺序（原始顺序）
            sorted_words = self.proper_words
        
        for word in sorted_words:
            self.term_listbox.insert(tk.END, word)
    
    def refresh_pending_list(self):
        """刷新导入预览列表"""
        self.pending_listbox.delete(0, tk.END)
        for wrong, right in self.import_preview:
            self.pending_listbox.insert(tk.END, f"{wrong} → {right}")
    
    def refresh_done_list(self):
        """刷新已通过列表"""
        self.done_listbox.delete(0, tk.END)
        
        # 先收集所有词条
        entries = []
        if os.path.exists(self.word_library_path):
            with open(self.word_library_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and '=' in line and not line.startswith('#'):
                        parts = line.split('=')
                        if len(parts) == 2:
                            wrong = parts[0].strip()
                            right = parts[1].strip()
                            entries.append((wrong, right))
        
        # 根据排序方式显示
        if self.done_sort_var.get() == "alpha":
            # 字母顺序排序（按错词排序）
            entries = sorted(entries, key=lambda x: x[0].lower())
        
        # 显示
        for wrong, right in entries:
            self.done_listbox.insert(tk.END, f"{wrong} → {right}")
    
    def refresh_view(self):
        """刷新查看词库"""
        self.view_text.config(state=tk.NORMAL)
        self.view_text.delete(1.0, tk.END)
        
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        if os.path.exists(self.word_library_md_path):
            with open(self.word_library_md_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
                # 替换时间显示，添加具体时分秒
                # 匹配 "创建时间：YYYY-MM-DD" 或 "创建时间：YYYY-MM-DD " 格式
                content = re.sub(
                    r'(创建时间：)(\d{4}-\d{2}-\d{2})\s*',
                    rf'\g<1>{current_time[:10]} {current_time[11:]}',
                    content
                )
                content = re.sub(
                    r'(最后更新：)(\d{4}-\d{2}-\d{2})\s*',
                    rf'\g<1>{current_time[:10]} {current_time[11:]}',
                    content
                )
                
                self.view_text.insert(1.0, content)
        
        self.view_text.config(state=tk.DISABLED)
    
    def on_term_double_click(self, event):
        """双击术语进行编辑"""
        selection = self.term_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        old_word = self.proper_words[idx]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑术语")
        dialog.geometry("400x160")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="输入新术语:", font=("Microsoft YaHei", 11)).pack(pady=10)
        
        entry = tk.Entry(dialog, font=("Microsoft YaHei", 11), width=40)
        entry.insert(0, old_word)
        entry.pack(pady=10)
        entry.select_range(0, tk.END)
        entry.focus()
        
        def on_ok():
            new_word = entry.get().strip()
            if new_word and new_word != old_word:
                self.proper_words[idx] = new_word
                self.save_proper_words()
                self.refresh_term_list()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        # 确定按钮 - 绿色背景白色文字，高度增加
        tk.Button(btn_frame, text="确定", command=on_ok, font=("Microsoft YaHei", 11), padx=25, pady=8,
                  bg="#4CAF50", fg="white", activebackground="#45a049", activeforeground="white",
                  width=10).pack(side=tk.LEFT, padx=10)
        # 取消按钮 - 灰色背景，高度增加
        tk.Button(btn_frame, text="取消", command=on_cancel, font=("Microsoft YaHei", 11), padx=25, pady=8,
                  bg="#9E9E9E", fg="white", activebackground="#757575", activeforeground="white",
                  width=10).pack(side=tk.LEFT, padx=10)
        
        dialog.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: on_cancel())
    
    def import_word_library(self):
        """手工导入词库"""
        filename = filedialog.askopenfilename(
            title="选择词库文件",
            filetypes=[
                ("文本文件", "*.txt"),
                ("Word文档", "*.docx"),
                ("所有文件", "*.*")
            ],
            initialdir=os.path.expanduser("~/Desktop")
        )
        
        if not filename:
            return
        
        # 解析词条
        entries = self.parse_word_file(filename)
        
        if not entries:
            messagebox.showinfo("提示", "未能从文件中解析到词条\n\n支持的格式：\n被替换词：正确的词\n被替换词 正确的词")
            return
        
        # 保存到预览
        self.import_preview = entries
        
        # 启用更新按钮
        self.update_btn.config(state=tk.NORMAL)
        
        # 切换到预览标签
        self.switch_tab("pending")
        
        # 更新状态
        self.status_label.config(text=f"专业术语: {len(self.proper_words)} 条 | 导入预览: {len(self.import_preview)} 条")
        
        messagebox.showinfo("导入成功", f"已解析 {len(entries)} 个词条\n\n请在「导入预览」标签中查看\n确认无误后点击「更新导入词库」")
    
    def parse_word_file(self, filepath):
        """解析词库文件"""
        entries = []
        
        try:
            # 读取文件
            if filepath.endswith('.docx'):
                if not DOCX_AVAILABLE:
                    messagebox.showerror("错误", "python-docx 未安装")
                    return entries
                
                doc = docx.Document(filepath)
                lines = [p.text for p in doc.paragraphs if p.text.strip()]
            else:
                with open(filepath, 'r', encoding='utf-8') as f:
                    lines = [line.strip() for line in f if line.strip()]
            
            # 解析每一行
            for line in lines:
                # 跳过注释行
                if line.startswith('#') or line.startswith('//'):
                    continue
                
                wrong = ""
                right = ""
                
                # 尝试不同的分隔符
                if '：' in line:  # 中文冒号
                    parts = line.split('：')
                    if len(parts) >= 2:
                        wrong = parts[0].strip()
                        right = parts[1].strip()
                elif ':' in line:  # 英文冒号
                    parts = line.split(':')
                    if len(parts) >= 2:
                        wrong = parts[0].strip()
                        right = parts[1].strip()
                elif '→' in line:  # 箭头
                    parts = line.split('→')
                    if len(parts) >= 2:
                        wrong = parts[0].strip()
                        right = parts[1].strip()
                elif '=' in line:  # 等号
                    parts = line.split('=')
                    if len(parts) >= 2:
                        wrong = parts[0].strip()
                        right = parts[1].strip()
                elif ' ' in line:  # 空格分隔（最后一个空格作为分隔）
                    # 找到最后一个空格
                    idx = line.rfind(' ')
                    if idx > 0:
                        wrong = line[:idx].strip()
                        right = line[idx+1:].strip()
                
                # 如果成功解析
                if wrong and right:
                    entries.append([wrong, right])
        
        except Exception as e:
            messagebox.showerror("错误", f"解析文件失败: {str(e)}")
        
        return entries
    
    def update_word_library(self):
        """更新导入词库"""
        if not self.import_preview:
            messagebox.showwarning("警告", "没有待导入的词条")
            return
        
        # 获取选中的条目
        selected = self.pending_listbox.curselection()
        
        entries_to_add = []
        if selected:
            # 只导入选中的
            for i in selected:
                entries_to_add.append(self.import_preview[i])
        else:
            # 导入全部
            entries_to_add = self.import_preview
        
        if not entries_to_add:
            messagebox.showwarning("警告", "请选择要导入的词条")
            return
        
        # 保存到词库（自动去重）
        count = self.save_word_library(entries_to_add)
        
        # 清空预览
        self.import_preview = []
        self.update_btn.config(state=tk.DISABLED)
        
        # 刷新
        self.switch_tab("done")
        self.status_label.config(text=f"专业术语: {len(self.proper_words)} 条 | 导入预览: 0 条")
        
        messagebox.showinfo("完成", f"已成功导入 {count} 个词条到词库（已自动过滤重复项）")
    
    def add_proper_word(self):
        """添加专业术语"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加专业术语")
        dialog.geometry("500x380")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 分隔符选择
        sep_frame = tk.Frame(dialog)
        sep_frame.pack(fill=tk.X, padx=15, pady=10)
        
        tk.Label(sep_frame, text="分隔符:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        
        sep_var = tk.StringVar(value="space")
        
        tk.Radiobutton(sep_frame, text="空格", variable=sep_var, value="space", 
                       font=("Microsoft YaHei", 9), command=lambda: parse_and_refresh()).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(sep_frame, text="逗号", variable=sep_var, value="comma", 
                       font=("Microsoft YaHei", 9), command=lambda: parse_and_refresh()).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(sep_frame, text="顿号", variable=sep_var, value="dunhao", 
                       font=("Microsoft YaHei", 9), command=lambda: parse_and_refresh()).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(sep_frame, text="换行", variable=sep_var, value="newline", 
                       font=("Microsoft YaHei", 9), command=lambda: parse_and_refresh()).pack(side=tk.LEFT, padx=5)
        
        # 输入说明
        tk.Label(dialog, text="批量输入（每行一个，或使用所选分隔符分隔）:", font=("Microsoft YaHei", 10)).pack(anchor=tk.W, padx=15)
        
        # 批量输入文本框
        batch_text = tk.Text(dialog, font=("Microsoft YaHei", 10), height=10, wrap=tk.WORD)
        batch_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        # 预览标签
        preview_label = tk.Label(dialog, text="", font=("Microsoft YaHei", 9), fg="gray")
        preview_label.pack(anchor=tk.W, padx=15)
        
        # 当前解析结果
        current_new_items = []
        
        def parse_and_refresh():
            """解析批量输入"""
            nonlocal current_new_items
            text = batch_text.get("1.0", tk.END).strip()
            if not text:
                preview_label.config(text="", fg="gray")
                current_new_items = []
                return
            
            # 根据分隔符分割
            if sep_var.get() == "space":
                items = text.split()
            elif sep_var.get() == "comma":
                items = text.replace('，', ',').split(',')
            elif sep_var.get() == "dunhao":
                items = text.replace('、', '、').split('、')
            else:  # newline
                items = text.split('\n')
            
            # 过滤空项并去重，同时去除前后空格和特殊字符
            cleaned_items = []
            for item in items:
                item = item.strip()
                # 去除常见的不可见字符
                item = item.replace('\u200b', '').replace('\ufeff', '').replace('\u3000', '')
                item = item.strip()
                if item:
                    cleaned_items.append(item)
            
            items = list(dict.fromkeys(cleaned_items))  # 保持顺序去重
            
            # 统计（精确匹配）
            existing = [item for item in items if item in self.proper_words]
            new_items = [item for item in items if item not in self.proper_words]
            current_new_items = new_items
            
            if items:
                # 显示解析结果
                if existing:
                    exist_str = f" 已存在: {', '.join(existing[:5])}"
                    if len(existing) > 5:
                        exist_str += f" 等{len(existing)}个"
                else:
                    exist_str = ""
                preview_label.config(
                    text=f"解析: {len(items)} 条 | 新增: {len(new_items)} 条{exist_str}",
                    fg="#4CAF50" if new_items else "#FF9800"
                )
            else:
                preview_label.config(text="", fg="gray")
        
        # 绑定文本变化事件
        batch_text.bind('<<Modified>>', lambda e: parse_and_refresh())
        
        # 按钮框架
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=15)
        
        def on_add():
            """添加并继续"""
            nonlocal current_new_items
            new_items = current_new_items
            
            if not new_items:
                messagebox.showwarning("警告", "没有可添加的新术语")
                return
            
            # 添加到列表
            self.proper_words.extend(new_items)
            self.save_proper_words()
            self.refresh_term_list()
            self.status_label.config(text=f"专业术语: {len(self.proper_words)} 条 | 导入预览: {len(self.import_preview)} 条")
            
            # 清空输入框并刷新
            batch_text.delete("1.0", tk.END)
            current_new_items = []
            preview_label.config(text="", fg="gray")
            
            messagebox.showinfo("成功", f"已添加 {len(new_items)} 个术语，可继续输入...")
        
        def on_done():
            """完成并关闭"""
            dialog.destroy()
        
        # 添加按钮 - 绿色
        tk.Button(btn_frame, text="添加并继续", command=on_add, font=("Microsoft YaHei", 11), padx=20, pady=8,
                  bg="#4CAF50", fg="white", activebackground="#45a049", activeforeground="white",
                  width=12).pack(side=tk.LEFT, padx=10)
        # 完成按钮 - 灰色
        tk.Button(btn_frame, text="完成", command=on_done, font=("Microsoft YaHei", 11), padx=20, pady=8,
                  bg="#9E9E9E", fg="white", activebackground="#757575", activeforeground="white",
                  width=10).pack(side=tk.LEFT, padx=10)
        
        dialog.bind('<Escape>', lambda e: on_done())
    
    def select_all(self):
        """全选"""
        if self.current_tab == "terms":
            self.term_listbox.select_set(0, tk.END)
            # 刷新显示以确保选中状态可见
            self.root.update()
        elif self.current_tab == "pending":
            self.pending_listbox.select_set(0, tk.END)
            self.root.update()
        elif self.current_tab == "review":
            # 全选待审核列表：将所有项标记为选中
            self.review_items = [(w, r, True) for w, r, s in self.review_items]
            self.refresh_review_list()
            self.update_status_label()
        elif self.current_tab == "done":
            self.done_listbox.select_set(0, tk.END)
            self.root.update()
    
    def deselect_all(self):
        """取消全选"""
        if self.current_tab == "terms":
            self.term_listbox.select_clear(0, tk.END)
            self.root.update()
        elif self.current_tab == "pending":
            self.pending_listbox.select_clear(0, tk.END)
            self.root.update()
        elif self.current_tab == "review":
            # 取消全选待审核列表：将所有项标记为未选中
            self.review_items = [(w, r, False) for w, r, s in self.review_items]
            self.refresh_review_list()
            self.update_status_label()
        elif self.current_tab == "done":
            self.done_listbox.select_clear(0, tk.END)
            self.root.update()
    
    def delete_selected(self):
        """删除选中项"""
        if self.current_tab == "terms":
            selected = self.term_listbox.curselection()
            if not selected:
                messagebox.showwarning("警告", "请先选择要删除的术语")
                return
            
            # 获取排序后的列表
            if self.sort_var.get() == "alpha":
                display_words = sorted(self.proper_words, key=lambda x: x.lower())
            else:
                display_words = self.proper_words
            
            # 删除选中的项（基于显示顺序）
            indices_to_delete = set(selected)
            new_words = [w for i, w in enumerate(display_words) if i not in indices_to_delete]
            self.proper_words = new_words
            
            self.save_proper_words()
            self.refresh_term_list()
            
        elif self.current_tab == "pending":
            selected = self.pending_listbox.curselection()
            if not selected:
                messagebox.showwarning("警告", "请先选择要删除的项")
                return
            
            for i in reversed(selected):
                del self.import_preview[i]
            
            self.refresh_pending_list()
            
            if not self.import_preview:
                self.update_btn.config(state=tk.DISABLED)
        
        elif self.current_tab == "review":
            # 对于待审核列表，直接删除所有标记为选中的项（点击列表项切换☑/☐）
            selected_count = sum(1 for w, r, s in self.review_items if s)
            if selected_count == 0:
                messagebox.showwarning("警告", "请先勾选要删除的项（点击列表项切换☑/☐）")
                return
            
            # 删除所有标记为选中的项
            self.review_items = [(w, r, s) for w, r, s in self.review_items if not s]
            self.refresh_review_list()
        
        self.update_status_label()
    
    def update_status_label(self):
        """更新状态栏"""
        selected_count = sum(1 for w, r, s in self.review_items if s)
        self.status_label.config(
            text=f"专业术语: {len(self.proper_words)} 条 | 导入预览: {len(self.import_preview)} 条 | 待审核: {len(self.review_items)} 条 | 已选: {selected_count} 条"
        )
    
    def batch_import_to_library(self):
        """批量导入待审核词条到词库"""
        if not self.review_items:
            messagebox.showwarning("警告", "没有待审核的词条")
            return
        
        # 获取选中的词条
        selected_items = [(w, r) for w, r, s in self.review_items if s]
        
        if not selected_items:
            messagebox.showwarning("警告", "请先勾选要导入的词条")
            return
        
        # 保存到词库（自动去重）
        count = self.save_word_library(selected_items)
        
        # 清空已导入的词条
        self.review_items = [(w, r, s) for w, r, s in self.review_items if not s]
        self.refresh_review_list()
        
        self.update_status_label()
        
        messagebox.showinfo("完成", f"已成功导入 {count} 个词条到词库（已自动过滤重复项）")


if __name__ == "__main__":
    app = ReadLibraryApp()

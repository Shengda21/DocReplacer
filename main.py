# -*- coding: utf-8 -*-
"""
文档词汇批量替换工具 - GUI 主程序
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# 拖放支持
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

try:
    import ttkbootstrap as ttkb
    from ttkbootstrap.constants import *
    USE_BOOTSTRAP = True
except ImportError:
    USE_BOOTSTRAP = False

from replacer import load_rules, replace_in_file, backup_file, restore_file, has_backup, SUPPORTED_EXTENSIONS

def resource_path(relative_path):
    """ 获取资源的绝对路径，兼容 PyInstaller 打包运行环境 """
    try:
        # PyInstaller 创建临时文件夹，将路径存入 _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ─── 常量 ──────────────────────────────────────────────────

APP_TITLE = "文档词汇批量替换工具"
WINDOW_SIZE = "960x680"
MIN_WIDTH = 800
MIN_HEIGHT = 560


# ─── 主应用类 ──────────────────────────────────────────────

class ReplacerApp:
    def __init__(self):
        # 创建主窗口（带拖放支持）
        if HAS_DND:
            if USE_BOOTSTRAP:
                # 兼容 ttkbootstrap，让它的 Window 继承 TkinterDnD.Tk
                class DnDWindow(ttkb.Window, TkinterDnD.DnDWrapper):
                    def __init__(self, *args, **kwargs):
                        ttkb.Window.__init__(self, *args, **kwargs)
                        self.TkdndVersion = TkinterDnD._require(self)
                self.root = DnDWindow(
                    title=APP_TITLE,
                    themename="cosmo",
                    size=(960, 680),
                    minsize=(MIN_WIDTH, MIN_HEIGHT),
                )
            else:
                self.root = TkinterDnD.Tk()
                self.root.title(APP_TITLE)
                self.root.geometry(WINDOW_SIZE)
                self.root.minsize(MIN_WIDTH, MIN_HEIGHT)
        else:
            if USE_BOOTSTRAP:
                self.root = ttkb.Window(
                    title=APP_TITLE,
                    themename="cosmo",
                    size=(960, 680),
                    minsize=(MIN_WIDTH, MIN_HEIGHT),
                )
            else:
                self.root = tk.Tk()
                self.root.title(APP_TITLE)
                self.root.geometry(WINDOW_SIZE)
                self.root.minsize(MIN_WIDTH, MIN_HEIGHT)

        # 设置图标
        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        # 数据
        self.rules = []            # [(原词, 替换词), ...]
        self.doc_files = []        # [filepath, ...]
        self.rules_filepath = ""   # 当前规则文件路径

        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"+{x}+{y}")

    def _build_ui(self):
        """构建整个 UI"""
        # 主容器
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ── 顶部工具栏 ──
        self._build_toolbar(main_frame)

        # ── 中间区域：规则表格 + 文件列表 ──
        content = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        content.pack(fill=tk.BOTH, expand=True, pady=(10, 5))

        # 左侧：替换规则表格
        self._build_rules_panel(content)

        # 右侧：文件列表
        self._build_files_panel(content)

        # ── 底部状态栏 ──
        self._build_statusbar(main_frame)

    def _build_toolbar(self, parent):
        """工具栏"""
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill=tk.X)

        # 使用样式化按钮
        btn_style = "primary.TButton" if USE_BOOTSTRAP else "TButton"
        warn_style = "warning.TButton" if USE_BOOTSTRAP else "TButton"
        success_style = "success.TButton" if USE_BOOTSTRAP else "TButton"
        danger_style = "danger.TButton" if USE_BOOTSTRAP else "TButton"

        self.btn_import_rules = ttk.Button(
            toolbar, text="📋 导入替换规则", style=btn_style,
            command=self._on_import_rules
        )
        self.btn_import_rules.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_export_rules = ttk.Button(
            toolbar, text="💾 导出替换规则", style=btn_style,
            command=self._on_export_rules
        )
        self.btn_export_rules.pack(side=tk.LEFT, padx=5)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        self.btn_replace = ttk.Button(
            toolbar, text="🔄 批量替换", style=success_style,
            command=self._on_replace
        )
        self.btn_replace.pack(side=tk.LEFT, padx=5)
        self.btn_replace.state(["disabled"])

        self.btn_restore = ttk.Button(
            toolbar, text="↩️ 还原", style=danger_style,
            command=self._on_restore
        )
        self.btn_restore.pack(side=tk.LEFT, padx=5)
        self.btn_restore.state(["disabled"])

        # 右侧：清空按钮
        self.btn_clear = ttk.Button(
            toolbar, text="🗑️ 清空", style=warn_style,
            command=self._on_clear
        )
        self.btn_clear.pack(side=tk.RIGHT, padx=(5, 0))

    def _build_rules_panel(self, parent):
        """替换规则表格面板"""
        frame = ttk.LabelFrame(parent if not isinstance(parent, ttk.PanedWindow) else None,
                               text="替换规则", padding=5)

        if isinstance(parent, ttk.PanedWindow):
            frame = ttk.Frame(parent, padding=0)
            parent.add(frame, weight=3)

            lbl = ttk.Label(frame, text="📋 替换规则", font=("Microsoft YaHei UI", 11, "bold"))
            lbl.pack(anchor=tk.W, pady=(0, 5))

        # 表格区域
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("index", "original", "replacement")
        self.rules_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", height=15
        )
        self.rules_tree.heading("index", text="序号")
        self.rules_tree.heading("original", text="原词")
        self.rules_tree.heading("replacement", text="替换词")
        self.rules_tree.column("index", width=50, minwidth=40, anchor=tk.CENTER)
        self.rules_tree.column("original", width=200, minwidth=100)
        self.rules_tree.column("replacement", width=200, minwidth=100)

        # 双击编辑
        self.rules_tree.bind("<Double-1>", self._on_edit_rule)

        # 滚动条
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.rules_tree.yview)
        self.rules_tree.configure(yscrollcommand=scrollbar.set)

        self.rules_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 注册拖放事件
        if HAS_DND:
            self.rules_tree.drop_target_register(DND_FILES)
            self.rules_tree.dnd_bind('<<Drop>>', self._on_drop_rules)

        # 规则操作按钮行
        rule_btn_frame = ttk.Frame(frame)
        rule_btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(rule_btn_frame, text="➕ 添加规则",
                   command=self._on_add_rule).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(rule_btn_frame, text="✏️ 编辑规则",
                   command=self._on_edit_rule).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(rule_btn_frame, text="➖ 删除规则",
                   command=self._on_delete_rule).pack(side=tk.LEFT)

        # 规则计数标签
        self.rules_count_var = tk.StringVar(value="共 0 条规则")
        ttk.Label(frame, textvariable=self.rules_count_var).pack(anchor=tk.W, pady=(3, 0))

    def _build_files_panel(self, parent):
        """文件列表面板"""
        if isinstance(parent, ttk.PanedWindow):
            frame = ttk.Frame(parent, padding=0)
            parent.add(frame, weight=2)

            lbl = ttk.Label(frame, text="📂 待处理文档", font=("Microsoft YaHei UI", 11, "bold"))
            lbl.pack(anchor=tk.W, pady=(0, 5))
        else:
            frame = ttk.LabelFrame(parent, text="待处理文档", padding=5)
            frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # 文件列表框
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.files_listbox = tk.Listbox(
            list_frame, selectmode=tk.EXTENDED,
            font=("Microsoft YaHei UI", 9)
        )
        files_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                                         command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)

        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        files_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 注册拖放事件
        if HAS_DND:
            self.files_listbox.drop_target_register(DND_FILES)
            self.files_listbox.dnd_bind('<<Drop>>', self._on_drop_docs)

        # 文件操作按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(btn_frame, text="➕ 添加文件",
                   command=self._on_select_docs).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="➖ 移除选中",
                   command=self._on_remove_selected_docs).pack(side=tk.LEFT)

        self.files_count_var = tk.StringVar(value="共 0 个文档")
        ttk.Label(frame, textvariable=self.files_count_var).pack(anchor=tk.W, pady=(3, 0))

        # 支持的文件类型小字提示
        supported_text = "支持格式: " + " ".join(SUPPORTED_EXTENSIONS)
        ttk.Label(
            frame, text=supported_text,
            font=("Microsoft YaHei UI", 7),
            foreground="#888888"
        ).pack(anchor=tk.W, pady=(2, 0))

    def _build_statusbar(self, parent):
        """状态栏"""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(5, 0))

        self.status_var = tk.StringVar(value="就绪 — 请先导入替换规则文件")
        status_label = ttk.Label(
            status_frame, textvariable=self.status_var,
            font=("Microsoft YaHei UI", 9), anchor=tk.W
        )
        status_label.pack(fill=tk.X)

        # 进度条
        self.progress = ttk.Progressbar(status_frame, mode="determinate", length=200)
        self.progress.pack(fill=tk.X, pady=(3, 0))

    # ─── 事件处理 ──────────────────────────────────────────

    def _on_import_rules(self):
        """导入替换规则文件"""
        filepath = filedialog.askopenfilename(
            title="选择替换规则文件",
            filetypes=[
                ("Excel 文件", "*.xlsx"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ]
        )
        if not filepath:
            return
        self._load_rules_from_file(filepath)

    def _on_drop_rules(self, event):
        """处理拖入替换规则文件"""
        # 可能有多个文件，或带有大括号 { C:\... }，解析出一个文件路径
        files = self.root.tk.splitlist(event.data)
        if files:
            filepath = files[0]
            if Path(filepath).suffix.lower() in [".xlsx", ".csv"]:
                self._load_rules_from_file(filepath)
            else:
                messagebox.showwarning("提示", "规则文件只支持 .xlsx 或 .csv 格式")

    def _load_rules_from_file(self, filepath):
        try:
            self.rules = load_rules(filepath)
            self.rules_filepath = filepath
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取规则文件：\n{e}")
            return

        if not self.rules:
            messagebox.showwarning("提示", "规则文件中未找到有效的替换规则。\n请确认文件包含两列数据。")
            return

        # 刷新表格
        self._refresh_rules_table()
        self._update_button_states()
        self.status_var.set(f"已导入 {len(self.rules)} 条替换规则  ←  {Path(filepath).name}")

    def _on_export_rules(self):
        """导出当前替换规则到文件"""
        if not self.rules:
            messagebox.showinfo("提示", "当前没有替换规则可导出。")
            return

        filepath = filedialog.asksaveasfilename(
            title="导出替换规则",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel 文件", "*.xlsx"),
                ("CSV 文件", "*.csv"),
            ]
        )
        if not filepath:
            return

        try:
            ext = Path(filepath).suffix.lower()
            if ext == ".xlsx":
                from openpyxl import Workbook
                wb = Workbook()
                ws = wb.active
                ws.title = "替换规则"
                ws.append(["原词", "替换词"])
                for old, new in self.rules:
                    ws.append([old, new])
                ws.column_dimensions['A'].width = 20
                ws.column_dimensions['B'].width = 20
                wb.save(filepath)
            elif ext == ".csv":
                import csv
                with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(["原词", "替换词"])
                    for old, new in self.rules:
                        writer.writerow([old, new])
            else:
                messagebox.showerror("错误", f"不支持的格式: {ext}")
                return

            messagebox.showinfo("导出成功", f"已导出 {len(self.rules)} 条规则到：\n{Path(filepath).name}")
            self.status_var.set(f"已导出 {len(self.rules)} 条规则 → {Path(filepath).name}")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出时出错：\n{e}")

    def _refresh_rules_table(self):
        """刷新替换规则表格"""
        self.rules_tree.delete(*self.rules_tree.get_children())
        for i, (old, new) in enumerate(self.rules, 1):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.rules_tree.insert("", tk.END, values=(i, old, new), tags=(tag,))

        # 设置交替行颜色
        self.rules_tree.tag_configure("evenrow", background="#f0f4f8")
        self.rules_tree.tag_configure("oddrow", background="#ffffff")

        self.rules_count_var.set(f"共 {len(self.rules)} 条规则")

    # ─── 手动编辑规则 ────────────────────────────────────────

    def _on_add_rule(self):
        """弹出对话框手动添加一条规则"""
        result = self._show_rule_dialog("添加替换规则", "", "")
        if result:
            old_text, new_text = result
            self.rules.append((old_text, new_text))
            self._refresh_rules_table()
            self._update_button_states()
            self.status_var.set(f"已添加规则: {old_text} → {new_text}")

    def _on_edit_rule(self, event=None):
        """编辑选中的规则"""
        selected = self.rules_tree.selection()
        if not selected:
            if event is None:  # 来自按钮点击，非双击
                messagebox.showinfo("提示", "请先在表格中选中一条规则")
            return

        item = selected[0]
        values = self.rules_tree.item(item, "values")
        idx = int(values[0]) - 1  # 获取序号（1-based）转为索引

        result = self._show_rule_dialog("编辑替换规则", values[1], values[2])
        if result:
            old_text, new_text = result
            self.rules[idx] = (old_text, new_text)
            self._refresh_rules_table()
            self.status_var.set(f"已更新规则: {old_text} → {new_text}")

    def _on_delete_rule(self):
        """删除选中的规则"""
        selected = self.rules_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先在表格中选中要删除的规则")
            return

        # 收集要删除的索引（从大到小排序以避免偏移）
        indices = []
        for item in selected:
            values = self.rules_tree.item(item, "values")
            indices.append(int(values[0]) - 1)
        indices.sort(reverse=True)

        for idx in indices:
            del self.rules[idx]

        self._refresh_rules_table()
        self._update_button_states()
        self.status_var.set(f"已删除 {len(indices)} 条规则")

    def _show_rule_dialog(self, title, old_val, new_val):
        """显示添加/编辑规则的弹窗，返回 (原词, 替换词) 或 None"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 90
        dialog.geometry(f"+{x}+{y}")

        result = [None]

        ttk.Label(dialog, text="原词:", font=("Microsoft YaHei UI", 10)).place(x=30, y=20)
        entry_old = ttk.Entry(dialog, width=30, font=("Microsoft YaHei UI", 10))
        entry_old.place(x=100, y=20)
        entry_old.insert(0, old_val)

        ttk.Label(dialog, text="替换词:", font=("Microsoft YaHei UI", 10)).place(x=30, y=60)
        entry_new = ttk.Entry(dialog, width=30, font=("Microsoft YaHei UI", 10))
        entry_new.place(x=100, y=60)
        entry_new.insert(0, new_val)

        def on_ok():
            o = entry_old.get().strip()
            n = entry_new.get().strip()
            if not o:
                messagebox.showwarning("提示", "原词不能为空", parent=dialog)
                return
            if not n:
                messagebox.showwarning("提示", "替换词不能为空", parent=dialog)
                return
            result[0] = (o, n)
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.place(x=100, y=110)
        ttk.Button(btn_frame, text="确定", command=on_ok, width=10).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side=tk.LEFT)

        # 回车确认
        dialog.bind("<Return>", lambda e: on_ok())
        dialog.bind("<Escape>", lambda e: on_cancel())
        entry_old.focus_set()

        dialog.wait_window()
        return result[0]

    def _on_select_docs(self):
        """选择待替换的文档"""
        ext_pattern = " ".join(f"*{e}" for e in SUPPORTED_EXTENSIONS)
        filepaths = filedialog.askopenfilenames(
            title="选择文档",
            filetypes=[
                ("支持的文档", ext_pattern),
                ("Word 文档", "*.docx *.doc"),
                ("Excel 文档", "*.xlsx *.xls"),
                ("PowerPoint 文档", "*.pptx *.ppt"),
                ("文本/CSV 文件", "*.txt *.csv"),
                ("所有文件", "*.*"),
            ]
        )
        if not filepaths:
            return

        for fp in filepaths:
            if fp not in self.doc_files:
                self.doc_files.append(fp)

        self._refresh_files_list()
        self._update_button_states()
        self.status_var.set(f"已选择 {len(self.doc_files)} 个文档")

    def _on_drop_docs(self, event):
        """处理拖入文档文件"""
        files = self.root.tk.splitlist(event.data)
        added = 0
        for filepath in files:
            if Path(filepath).suffix.lower() in SUPPORTED_EXTENSIONS:
                if filepath not in self.doc_files:
                    self.doc_files.append(filepath)
                    added += 1
        
        if added > 0:
            self._refresh_files_list()
            self._update_button_states()
            self.status_var.set(f"已选择 {len(self.doc_files)} 个文档 (新增 {added} 个)")
        else:
            if files:  # 如果拖入了文件但是没加进去，说明格式不对
                supported = " ".join(SUPPORTED_EXTENSIONS)
                messagebox.showwarning("提示", f"支持的格式: {supported}")

    def _refresh_files_list(self):
        """刷新文件列表"""
        self.files_listbox.delete(0, tk.END)
        for fp in self.doc_files:
            name = Path(fp).name
            backup_mark = " ✅有备份" if has_backup(fp) else ""
            self.files_listbox.insert(tk.END, f"{name}{backup_mark}")
        self.files_count_var.set(f"共 {len(self.doc_files)} 个文档")

    def _on_remove_selected_docs(self):
        """移除选中的文档"""
        selected = self.files_listbox.curselection()
        if not selected:
            return
        # 从后往前删除以避免索引偏移
        for i in reversed(selected):
            del self.doc_files[i]
        self._refresh_files_list()
        self._update_button_states()

    def _on_replace(self):
        """批量替换"""
        if not self.rules:
            messagebox.showwarning("提示", "请先导入替换规则文件！")
            return
        if not self.doc_files:
            messagebox.showwarning("提示", "请先选择待替换的文档！")
            return

        # 确认对话框
        msg = (
            f"即将对 {len(self.doc_files)} 个文档执行 {len(self.rules)} 条替换规则。\n\n"
            f"替换前将自动备份原始文件到 _backup 文件夹。\n\n"
            f"是否继续？"
        )
        if not messagebox.askyesno("确认替换", msg):
            return

        # 在后台线程执行替换
        self._set_buttons_enabled(False)
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.doc_files)

        thread = threading.Thread(target=self._do_replace, daemon=True)
        thread.start()

    def _do_replace(self):
        """后台执行替换"""
        total_replaced = 0
        errors = []

        for i, fp in enumerate(self.doc_files):
            try:
                self.root.after(0, self.status_var.set,
                                f"正在处理 ({i+1}/{len(self.doc_files)}): {Path(fp).name}")

                # 备份
                backup_file(fp)

                # 替换
                result = replace_in_file(fp, self.rules)
                total_replaced += result["total_replacements"]

            except Exception as e:
                errors.append(f"{Path(fp).name}: {e}")

            self.root.after(0, self._update_progress, i + 1)

        # 完成
        self.root.after(0, self._replace_done, total_replaced, errors)

    def _update_progress(self, value):
        self.progress["value"] = value

    def _replace_done(self, total_replaced, errors):
        """替换完成回调"""
        self._set_buttons_enabled(True)
        self._refresh_files_list()

        if errors:
            error_msg = "\n".join(errors)
            messagebox.showwarning(
                "替换完成（有错误）",
                f"替换完成，共替换 {total_replaced} 处。\n\n"
                f"以下文件出现错误：\n{error_msg}"
            )
        else:
            messagebox.showinfo(
                "替换完成",
                f"全部完成！共替换 {total_replaced} 处。\n"
                f"原始文件已备份到 _backup 文件夹。"
            )
        self.status_var.set(f"替换完成 — 共替换 {total_replaced} 处")

    def _on_restore(self):
        """还原所有文档"""
        # 检查哪些文件有备份
        restorable = [fp for fp in self.doc_files if has_backup(fp)]
        if not restorable:
            messagebox.showinfo("提示", "没有可还原的文件。\n请确认已执行过替换操作。")
            return

        msg = f"将还原 {len(restorable)} 个文件到替换前的状态。\n\n是否继续？"
        if not messagebox.askyesno("确认还原", msg):
            return

        success = 0
        errors = []
        for fp in restorable:
            try:
                if restore_file(fp):
                    success += 1
                else:
                    errors.append(f"{Path(fp).name}: 备份文件不存在")
            except Exception as e:
                errors.append(f"{Path(fp).name}: {e}")

        self._refresh_files_list()

        if errors:
            messagebox.showwarning(
                "还原完成（有错误）",
                f"成功还原 {success} 个文件。\n\n"
                f"以下文件出现错误：\n" + "\n".join(errors)
            )
        else:
            messagebox.showinfo("还原完成", f"已成功还原 {success} 个文件！")

        self.status_var.set(f"还原完成 — 成功还原 {success} 个文件")

    def _on_clear(self):
        """清空所有数据"""
        self.rules.clear()
        self.doc_files.clear()
        self.rules_filepath = ""
        self.rules_tree.delete(*self.rules_tree.get_children())
        self.files_listbox.delete(0, tk.END)
        self.rules_count_var.set("共 0 条规则")
        self.files_count_var.set("共 0 个文档")
        self._update_button_states()
        self.status_var.set("就绪 — 请先导入替换规则文件")
        self.progress["value"] = 0

    def _update_button_states(self):
        """更新按钮启用/禁用状态"""
        has_rules = len(self.rules) > 0
        has_docs = len(self.doc_files) > 0
        has_restorable = any(has_backup(fp) for fp in self.doc_files) if has_docs else False

        if has_rules and has_docs:
            self.btn_replace.state(["!disabled"])
        else:
            self.btn_replace.state(["disabled"])

        if has_restorable:
            self.btn_restore.state(["!disabled"])
        else:
            self.btn_restore.state(["disabled"])

    def _set_buttons_enabled(self, enabled: bool):
        """批量设置按钮可用状态"""
        state = ["!disabled"] if enabled else ["disabled"]
        self.btn_import_rules.state(state)
        self.btn_export_rules.state(state)
        self.btn_replace.state(state)
        self.btn_restore.state(state)
        self.btn_clear.state(state)

    def run(self):
        self.root.mainloop()


# ─── 入口 ────────────────────────────────────────────────

if __name__ == "__main__":
    app = ReplacerApp()
    app.run()

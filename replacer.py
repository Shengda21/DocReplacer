# -*- coding: utf-8 -*-
"""
文档词汇批量替换 - 核心逻辑模块
"""

import os
import csv
import shutil
import re
from pathlib import Path
from typing import List, Tuple

from docx import Document
from openpyxl import load_workbook


# ─── 替换规则加载 ────────────────────────────────────────────

def load_rules(filepath: str) -> List[Tuple[str, str]]:
    """
    从 xlsx 或 csv 文件加载替换规则。
    返回 [(原词, 替换词), ...] 列表。
    """
    ext = Path(filepath).suffix.lower()
    if ext == ".xlsx":
        return _load_rules_xlsx(filepath)
    elif ext == ".csv":
        return _load_rules_csv(filepath)
    else:
        raise ValueError(f"不支持的文件格式: {ext}，请使用 .xlsx 或 .csv 文件")


def _load_rules_xlsx(filepath: str) -> List[Tuple[str, str]]:
    wb = load_workbook(filepath, read_only=True)
    ws = wb.active
    rules = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            # 跳过标题行（如果第一行看起来像标题）
            if row and isinstance(row[0], str) and row[0].strip() in ("原词", "原文", "查找", "原始"):
                continue
        if row and len(row) >= 2 and row[0] and row[1]:
            rules.append((str(row[0]).strip(), str(row[1]).strip()))
    wb.close()
    return rules


def _load_rules_csv(filepath: str) -> List[Tuple[str, str]]:
    rules = []
    # 尝试检测编码
    for encoding in ("utf-8-sig", "utf-8", "gbk", "gb2312", "latin1"):
        try:
            with open(filepath, "r", encoding=encoding) as f:
                reader = csv.reader(f)
                first = True
                for row in reader:
                    if first:
                        first = False
                        if row and row[0].strip() in ("原词", "原文", "查找", "原始"):
                            continue
                    if len(row) >= 2 and row[0].strip() and row[1].strip():
                        rules.append((row[0].strip(), row[1].strip()))
            return rules
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("无法读取 CSV 文件，请确认文件编码为 UTF-8 或 GBK")


# ─── 备份与还原 ─────────────────────────────────────────────

BACKUP_DIR_NAME = "_backup"


def _get_backup_path(filepath: str) -> str:
    """获取备份文件的路径"""
    p = Path(filepath)
    backup_dir = p.parent / BACKUP_DIR_NAME
    return str(backup_dir / p.name)


def backup_file(filepath: str) -> str:
    """
    备份原文件到 _backup 文件夹。
    返回备份文件路径。
    """
    p = Path(filepath)
    backup_dir = p.parent / BACKUP_DIR_NAME
    backup_dir.mkdir(exist_ok=True)
    backup_path = backup_dir / p.name
    shutil.copy2(filepath, backup_path)
    return str(backup_path)


def restore_file(filepath: str) -> bool:
    """
    从备份还原文件。
    返回是否成功还原。
    """
    backup_path = _get_backup_path(filepath)
    if not os.path.exists(backup_path):
        return False
    shutil.copy2(backup_path, filepath)
    return True


def has_backup(filepath: str) -> bool:
    """检查文件是否有备份"""
    return os.path.exists(_get_backup_path(filepath))


# ─── 文档替换核心 ─────────────────────────────────────────────

def replace_in_docx(filepath: str, rules: List[Tuple[str, str]]) -> dict:
    """
    对 docx 文件执行批量替换。
    保留原始格式，处理跨 run 的文本匹配。

    返回 {"total_replacements": int, "detail": {原词: 替换次数}}
    """
    doc = Document(filepath)
    detail = {}

    for old_text, new_text in rules:
        count = 0
        # 替换段落中的文本
        for para in doc.paragraphs:
            count += _replace_in_paragraph(para, old_text, new_text)

        # 替换表格中的文本
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        count += _replace_in_paragraph(para, old_text, new_text)

        # 替换页眉页脚中的文本
        for section in doc.sections:
            for header_footer in [section.header, section.footer]:
                if header_footer is not None:
                    for para in header_footer.paragraphs:
                        count += _replace_in_paragraph(para, old_text, new_text)
                    for table in header_footer.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for para in cell.paragraphs:
                                    count += _replace_in_paragraph(para, old_text, new_text)

        if count > 0:
            detail[old_text] = count

    doc.save(filepath)

    total = sum(detail.values())
    return {"total_replacements": total, "detail": detail}


def _replace_in_paragraph(paragraph, old_text: str, new_text: str) -> int:
    """
    在段落中替换文本，尽量保留原始格式。
    处理跨 run 的文本匹配。
    返回替换次数。
    """
    # 先尝试简单的单 run 替换
    count = 0
    for run in paragraph.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)
            count += run.text.count(new_text)  # 近似计数
            # 重新计算：用替换前后的差异来准确计数
    if count > 0:
        return count

    # 如果单 run 没找到，尝试跨 run 匹配
    full_text = "".join(run.text for run in paragraph.runs)
    if old_text not in full_text:
        return 0

    # 跨 run 替换策略
    count = full_text.count(old_text)
    if count == 0:
        return 0

    return _cross_run_replace(paragraph, old_text, new_text)


def _cross_run_replace(paragraph, old_text: str, new_text: str) -> int:
    """
    处理跨 run 的文本替换。
    核心策略：构建字符到 run 的映射，找到匹配位置后，
    在第一个涉及的 run 中放入替换文本，清空其余涉及的 run 中对应的字符。
    """
    runs = paragraph.runs
    if not runs:
        return 0

    # 构建字符位置到 (run_index, char_index_in_run) 的映射
    char_map = []
    for run_idx, run in enumerate(runs):
        for char_idx in range(len(run.text)):
            char_map.append((run_idx, char_idx))

    full_text = "".join(run.text for run in runs)
    count = 0
    search_start = 0

    while True:
        pos = full_text.find(old_text, search_start)
        if pos == -1:
            break

        count += 1
        match_end = pos + len(old_text)

        # 找到匹配涉及的 run 范围
        first_run_idx = char_map[pos][0]
        last_run_idx = char_map[match_end - 1][0]

        # 在每个涉及的 run 中处理文本
        for run_idx in range(first_run_idx, last_run_idx + 1):
            run = runs[run_idx]
            run_text = run.text

            # 计算当前 run 中需要替换的字符范围
            # run 在全文中的起始位置
            run_start_in_full = sum(len(runs[i].text) for i in range(run_idx))
            run_end_in_full = run_start_in_full + len(run_text)

            # 匹配区域在当前 run 中的范围
            local_start = max(0, pos - run_start_in_full)
            local_end = min(len(run_text), match_end - run_start_in_full)

            if run_idx == first_run_idx:
                # 第一个 run：用替换文本替换匹配部分
                run.text = run_text[:local_start] + new_text + run_text[local_end:]
            else:
                # 后续 run：仅移除匹配部分
                run.text = run_text[:local_start] + run_text[local_end:]

        # 由于文本已变化，需要重建映射后继续（简单起见，递归处理剩余）
        # 这里简单 break 后再调用一次
        return count + _cross_run_replace(paragraph, old_text, new_text)

    return count

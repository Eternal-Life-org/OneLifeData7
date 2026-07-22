#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
translate_objects.py - 用 翻译表.xlsx 覆盖 objects/*.txt 的名称行(第2行)

用法:
    cd OneLifeData7
    python translate_objects.py              # 交互式选择模式
    python translate_objects.py --mode 1     # 1=翻译英文 2=翻译中文 3=追加英文(中文+英文)
    python translate_objects.py --dry-run     # 只扫描不覆盖(测试用)
    python translate_objects.py --ignore-errors  # 输出异常但仍执行覆盖(用xlsx修复object)

xlsx 结构 (Elife sheet):
    A=key(物品id)  B=English  C=Chinese  D=Label(词条, 自带 # 前缀)

三种模式构建的第2行:
    1 翻译英文:  English + Label
    2 翻译中文:  Chinese + Label
    3 追加英文:  Chinese + English + Label   (中英之间无空格)

名称与 Label 拼接时中间无空格(Label 自带 #, 直接拼接)。

扫描规则(覆盖前):
    - object 文件不存在                       -> warning (跳过, 不阻止覆盖)
    - xlsx 缺少该模式的翻译列                 -> warning (跳过该行不覆盖)
    - object 第2行为空                       -> 异常
    - xlsx Label 与 object 第2行现有 # 词条不一致 -> 异常
    有异常时默认退出不覆盖; 加 --ignore-errors 则输出异常后继续覆盖
    (用 xlsx 数据覆盖 object, 可修复 label 不一致/空行)。
"""

import sys
import os
import re
import zipfile
import xml.etree.ElementTree as ET

NS  = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# Elife sheet 名称
SHEET_NAME = 'Elife'
# 默认文件/目录 (相对运行时 cwd)
DEFAULT_XLSX = '翻译表.xlsx'
DEFAULT_OBJ_DIR = 'objects'


# ---------- xlsx 读取 (不依赖 openpyxl, 直接解析 XML) ----------

def _find_elife_sheet_path(z):
    """在 xlsx 里找 Elife sheet 对应的 worksheet xml 路径"""
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rid_target = {}
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    for rel in rels:
        rid_target[rel.get('Id')] = rel.get('Target')

    for sh in wb.findall(f'{NS}sheets/{NS}sheet'):
        if sh.get('name') == SHEET_NAME:
            rid = sh.get(f'{RNS}id')
            target = rid_target.get(rid)
            if target:
                if not target.startswith('/'):
                    target = 'xl/' + target
                else:
                    target = target.lstrip('/')
                return target
    return None


def _cell_value(c, shared):
    """取一个 <c> 单元格的值"""
    t = c.get('t')
    v = c.find(f'{NS}v')
    isn = c.find(f'{NS}is')
    if t == 's' and v is not None:
        return shared[int(v.text)]
    if isn is not None:
        return ''.join(tt.text or '' for tt in isn.iter(f'{NS}t'))
    if v is not None:
        return v.text or ''
    return ''


def clean_cell(s):
    """单元格清洗: 去内部回车换行, 去前后空白"""
    if s is None:
        return ''
    s = str(s)
    s = s.replace('\r', '').replace('\n', '')
    return s.strip()


def read_xlsx(xlsx_path):
    """读取 Elife sheet, 返回 dict[key] = {english, chinese, label}"""
    z = zipfile.ZipFile(xlsx_path)
    # sharedStrings
    ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
    shared = []
    for si in ss_tree.findall(f'{NS}si'):
        shared.append(''.join(t.text or '' for t in si.iter(f'{NS}t')))

    sheet_path = _find_elife_sheet_path(z)
    if sheet_path is None:
        raise RuntimeError(f"xlsx 里找不到名为 '{SHEET_NAME}' 的 sheet")

    sh_tree = ET.fromstring(z.read(sheet_path))
    rows = sh_tree.findall(f'{NS}sheetData/{NS}row')

    translations = {}
    for row in rows:
        cells = {}
        for c in row.findall(f'{NS}c'):
            ref = c.get('r')
            col = re.match(r'([A-Z]+)', ref).group(1)
            cells[col] = _cell_value(c, shared)
        key = clean_cell(cells.get('A', ''))
        # 只接受正整数 key (object id)
        if not re.match(r'^\d+$', key):
            continue
        translations[key] = {
            'english': clean_cell(cells.get('B', '')),
            'chinese': clean_cell(cells.get('C', '')),
            'label':   clean_cell(cells.get('D', '')),
        }
    return translations


# ---------- object 文件读写 ----------

def read_obj_line2(obj_path):
    """读 object 第2行(已去 \r\n); 文件不存在或不足2行返回 None"""
    try:
        with open(obj_path, encoding='utf-8') as f:
            content = f.read()
    except FileNotFoundError:
        return None
    content = content.replace('\r\n', '\n').replace('\r', '\n')
    lines = content.split('\n')
    if len(lines) < 2:
        return None
    line2 = lines[1]
    if line2.strip() == '':
        return ''
    return line2


def write_obj_line2(obj_path, new_line2):
    """覆盖 object 第2行, 保留其余行, 统一 LF"""
    with open(obj_path, encoding='utf-8') as f:
        content = f.read()
    content = content.replace('\r\n', '\n').replace('\r', '\n')
    lines = content.split('\n')
    if len(lines) < 2:
        return False
    lines[1] = new_line2
    with open(obj_path, 'w', encoding='utf-8', newline='') as f:
        f.write('\n'.join(lines))
    return True


# ---------- 名称构建 ----------

def mode_label(mode):
    return {1: '翻译英文', 2: '翻译中文', 3: '追加英文(中文+英文)'}[mode]


def mode_needs(mode):
    """该模式需要的翻译列 (用于扫描提示缺失)"""
    if mode == 1:
        return ['english']
    if mode == 2:
        return ['chinese']
    if mode == 3:
        return ['chinese', 'english']
    return []


def has_name(mode, t):
    """该模式下是否有可用名称 (决定是否跳过覆盖)
    mode 3 缺一列时用另一列 fallback, 都缺才跳过"""
    if mode == 1:
        return bool(t['english'])
    if mode == 2:
        return bool(t['chinese'])
    if mode == 3:
        return bool(t['chinese']) or bool(t['english'])
    return False


def build_name(mode, t):
    """构建名称部分(不含 label)"""
    if mode == 1:
        return t['english']
    if mode == 2:
        return t['chinese']
    if mode == 3:
        cn = t['chinese']
        en = t['english']
        if cn and en:
            # 中英文都有: 一致则保留一份, 否则拼接
            return cn if cn == en else cn + en
        # 只有其一: 用有的那个 (mode 3 对单语物品的 fallback)
        return cn or en
    return ''


def build_line2(mode, t):
    """构建完整第2行: 名称 + label (label 自带 #, 无空格)"""
    return build_name(mode, t) + t['label']


def extract_label(desc):
    """从描述行提取 # 后词条(含 #, 已 strip); 无 # 返回 ''"""
    idx = desc.find('#')
    if idx == -1:
        return ''
    return desc[idx:].strip()


# ---------- 扫描 ----------

def scan(translations, mode, obj_dir):
    """扫描所有 key, 返回 (errors, warnings)"""
    errors = []
    warnings = []
    needed = mode_needs(mode)

    for key in sorted(translations.keys(), key=lambda x: int(x)):
        t = translations[key]
        obj_path = os.path.join(obj_dir, f'{key}.txt')
        line2 = read_obj_line2(obj_path)

        # 1. object 文件不存在/行数不足 -> warning (xlsx 可有多余条目, 跳过即可)
        if line2 is None:
            warnings.append(f"key {key}: object 文件不存在或行数不足2行, 跳过")
            continue
        # object 第2行为空 -> 异常
        if line2 == '':
            errors.append(f"key {key}: object 第2行为空")
            continue

        # xlsx 中英文均为空 -> 异常 (无翻译数据, 不覆盖)
        if not t['english'] and not t['chinese']:
            errors.append(f"key {key}: xlsx 中英文均为空")
            continue

        # 2. xlsx 缺翻译列 -> warning (只缺一列, mode 3 可 fallback)
        missing = [c for c in needed if not t[c]]
        if missing:
            warnings.append(
                f"key {key}: xlsx 缺少 {','.join(missing)} 翻译")

        # 3. label 不一致 -> 异常
        obj_label = extract_label(line2)
        xlsx_label = t['label']
        if obj_label != xlsx_label:
            errors.append(
                f"key {key}: label 不一致  object={obj_label!r}  xlsx={xlsx_label!r}")

    return errors, warnings


# ---------- 覆盖 ----------

def translate(translations, mode, obj_dir):
    """执行覆盖, 返回成功覆盖的行数"""
    count = 0
    skipped = 0
    for key in sorted(translations.keys(), key=lambda x: int(x)):
        t = translations[key]
        # 无可用名称才跳过 (mode 3 缺一列时用另一列 fallback)
        if not has_name(mode, t):
            skipped += 1
            continue
        obj_path = os.path.join(obj_dir, f'{key}.txt')
        if not os.path.isfile(obj_path):
            skipped += 1
            continue
        if write_obj_line2(obj_path, build_line2(mode, t)):
            count += 1
        else:
            skipped += 1
    return count, skipped


# ---------- 交互 ----------

def ask_mode():
    print("请选择翻译模式:")
    print("  1: 翻译英文   (第2行 = English + Label)")
    print("  2: 翻译中文   (第2行 = Chinese + Label)")
    print("  3: 追加英文   (第2行 = Chinese + English + Label)")
    while True:
        try:
            m = int(input("请输入 1/2/3: ").strip())
            if m in (1, 2, 3):
                return m
        except ValueError:
            pass
        print("输入无效, 请输入 1/2/3")


# ---------- main ----------

def main():
    dry_run = '--dry-run' in sys.argv
    ignore_errors = '--ignore-errors' in sys.argv
    mode = None
    for a in sys.argv[1:]:
        if a.startswith('--mode='):
            mode = int(a.split('=', 1)[1])
        elif a == '--mode':
            pass
    # --mode N 形式
    if mode is None:
        args = [a for a in sys.argv[1:] if not a.startswith('--')]
        for a in args:
            if a in ('1', '2', '3'):
                mode = int(a)
                break
    if mode is None:
        mode = ask_mode()

    xlsx_path = DEFAULT_XLSX
    obj_dir = DEFAULT_OBJ_DIR

    print(f"模式: {mode_label(mode)}")
    print(f"读取 xlsx: {xlsx_path}")
    if not os.path.isfile(xlsx_path):
        print(f"错误: 找不到 xlsx 文件 '{xlsx_path}'")
        sys.exit(1)

    print("解析 xlsx ...")
    translations = read_xlsx(xlsx_path)
    print(f"xlsx 共 {len(translations)} 条 key")

    print("扫描 object 第2行 ...")
    errors, warnings = scan(translations, mode, obj_dir)

    # 输出 warning
    if warnings:
        print(f"\n--- warnings: {len(warnings)} 条 ---")
        for w in warnings:
            print(f"  [WARN] {w}")

    # 输出异常
    if errors:
        print(f"\n--- 异常: {len(errors)} 条 ---")
        for e in errors:
            print(f"  [ERROR] {e}")
        if ignore_errors:
            print(f"\n--ignore-errors: 忽略 {len(errors)} 处异常, 继续覆盖 "
                  f"(将用 xlsx 数据覆盖 object, 可修复 label 不一致/空行)。")
        else:
            print(f"\n扫描发现 {len(errors)} 处异常, 终止覆盖。"
                  f"请先修复, 或加 --ignore-errors 忽略异常继续覆盖。")
            sys.exit(1)

    print(f"\n扫描通过: {len(translations)} 条, warning {len(warnings)} 条, "
          f"异常 {len(errors)} 条"
          f"{' (已忽略)' if ignore_errors else ''}。")

    if dry_run:
        print("--dry-run 模式, 不执行覆盖。")
        return

    print("开始覆盖 ...")
    count, skipped = translate(translations, mode, obj_dir)
    print(f"完成: 覆盖 {count} 个 object, 跳过 {skipped} 个。")


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
translate.py - 用 翻译表.xlsx 覆盖 objects/*.txt 的名称行(第2行)

用法:
    cd OneLifeData7
    python translate.py              # 同步工具菜单: 新增物品/导入修改/应用翻译/撤回
    python translate.py --mode 1     # 直接执行: 1=翻译英文 2=翻译中文 3=追加英文(中文+英文)
    python translate.py --mode 3 --dry-run    # 只扫描不覆盖(测试用, 附带改动统计)
    python translate.py --mode 3 --ignore-errors  # 输出异常但仍执行覆盖
    python translate.py --check-transfer  # 只读: 检测"两边名字反过来"的疑似反转并用日志确认

菜单九个按钮(每次写入前自动备份, 可撤回一次):
    1. 新增物品 (本地 -> 翻译表): objects/ 里表上没有的编号, 追加为新行(B英文/C中文/D后缀)
    2. 导入修改 (本地 -> 翻译表): 本地第2行与表(实装B/C/D)不一致的, 用本地覆盖 B/C/D
    3. 应用翻译 (翻译表 -> 本地): 可选择 1 翻译英文 / 2 翻译中文 / 3 追加英文(中文+英文)
    4. 撤回: 回退到上一次操作前的状态
    8. 转移检测 (只读): 本地 objects 与表比对, 只有"两边都反过来"才算疑似反转
       (本地 X 与 Y 的名字在表里互相交换, 例: 本地 333=A 666=B, 表 333=B 666=A),
       单边撞名不算; 本地与表名字相同(仅后缀差异)的不比对。读取 idTransferLog.txt
       确认: 有对应记录=疑似反转确认, 无对应记录=按普通差异处理。
       日志查找: 环境变量 ID_TRANSFER_LOG 指定路径, 否则按
       当前目录/OneLifeData7/仓库根/output 顺序自动找。
    9. 删除旧翻译 (本地翻译表): 翻译表里有、本地 objects 没有的编号, 确认后删行。
       删除前先打印完整清单预览, y 才删, 其他键取消; 写前自动备份, 可菜单 4 撤回。
    备份在 OneLifeData7/sync_backup/

xlsx 结构 (Elife sheet, 按表头识别列):
    A=key(物品id)  B=English  C=简体中文（实装）  D=后缀（实装）
    E=英文（录入区）  F=简体中文（录入区）  G=后缀（录入区）  —— 录入区仅供调整, 本脚本不读取
    H=修改状态(公式, 只比名字)  I=备注

只有"实装"两列会被录入:
    C = 中文名(不含后缀)    D = 后缀(以#开头, 可为空)

后缀规范: 每个 # 后必须有一个空格(行业规范)。脚本读表时自动规范化,
'#Flowering' 与 '# Flowering' 视为同一后缀, 输出统一为 '# Flowering'。

三种模式构建的第2行:
    1 翻译英文:  English + 后缀
    2 翻译中文:  Chinese + 后缀
    3 追加英文:  Chinese + English + 后缀   (名字与后缀拼接均无额外空格)

扫描规则(覆盖前) —— 只检查名字, 不检查后缀:
    - object 文件不存在       -> warning (跳过, 不阻止覆盖)
    - xlsx 缺少该模式的翻译列 -> warning (跳过该行不覆盖)
    - object 第2行为空        -> 异常
    - xlsx 中英文均为空       -> 异常
    有异常时默认退出不覆盖; 加 --ignore-errors 则输出异常后继续覆盖
    (用 xlsx 数据覆盖 object, 可修复空行)。

兼容旧格式表(表头 key/English/Chinese/Label 或无表头的 B/C/D 固定列)。
"""

import sys
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
import openpyxl

NS  = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

# Elife sheet 名称
SHEET_NAME = 'Elife'
# 默认文件/目录 (相对运行时 cwd)
DEFAULT_XLSX = '翻译表.xlsx'
DEFAULT_OBJ_DIR = 'objects'

# 表格工具导出坏 styles.xml(如空 <fill/>)时 openpyxl 报这些错; 修复脚本见报错提示
LOAD_ERRORS = (TypeError, KeyError, ValueError)


def guard_load_workbook(**kwargs):
    """openpyxl 加载 + 坏样式容错; 失败时打印可操作的提示并返回 None"""
    try:
        return openpyxl.load_workbook(DEFAULT_XLSX, **kwargs)
    except LOAD_ERRORS as e:
        print(f'\n无法用 openpyxl 打开 {DEFAULT_XLSX}: {type(e).__name__}: {e}')
        print('   表格工具导出的 xlsx 可能含空 <fill/> 等坏样式(openpyxl 解析不了)。')
        print('   修复: 在 claude AI 目录执行')
        print('       python fix_xlsx_empty_fills.py <翻译表.xlsx>')
        print('   (只改 styles.xml, 单元格数据与格式不变, 自动备份; --check 可先预览)')
        print('   注意: 仅 --mode 1/2/3 走原生 XML 解析, 不受影响。')
        return None
    except OSError as e:
        print(f'\n无法读取 {DEFAULT_XLSX}: {e}')
        return None


# ---------- xlsx 读取 (不依赖 openpyxl, 直接解析 XML) ----------

def _load_shared(z):
    """sharedStrings 可选: openpyxl 等工具保存的表用内联字符串, 没有该文件"""
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join(t.text or '' for t in si.iter(f'{NS}t'))
            for si in ss_tree.findall(f'{NS}si')]


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
    """取一个 <c> 单元格的值(支持共享字符串/内联字符串/数字)"""
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


def _map_columns(header_cells):
    """按表头行识别列; 返回 (key, english, chinese, label) 列字母, 识别失败返回 None"""
    def find(pred):
        for col, v in header_cells.items():
            if pred(v.strip()):
                return col
        return None

    eng = find(lambda v: v.lower() == 'english' or ('实装' in v and '英文' in v))
    chn = find(lambda v: ('实装' in v and '中文' in v) or v.lower() == 'chinese')
    lab = find(lambda v: ('实装' in v and '后缀' in v) or v.lower() == 'label')
    key = find(lambda v: v.lower() == 'key')
    if eng and chn:
        return (key or 'A', eng, chn, lab)
    return None


def read_xlsx(xlsx_path):
    """读取 Elife sheet, 返回 dict[key] = {english, chinese, label}
    列按表头识别(实装列); 无表头或表头不认识时回退旧格式固定列 A/B/C/D"""
    z = zipfile.ZipFile(xlsx_path)
    shared = _load_shared(z)

    sheet_path = _find_elife_sheet_path(z)
    if sheet_path is None:
        raise RuntimeError(f"xlsx 里找不到名为 '{SHEET_NAME}' 的 sheet")

    sh_tree = ET.fromstring(z.read(sheet_path))
    xml_rows = sh_tree.findall(f'{NS}sheetData/{NS}row')

    translations = {}
    cols = None
    for row in xml_rows:
        cells = {}
        for c in row.findall(f'{NS}c'):
            ref = c.get('r')
            col = re.match(r'([A-Z]+)', ref).group(1)
            cells[col] = _cell_value(c, shared)

        if cols is None:
            # 首行: 是表头则识别列, 是数据(key为纯数字)则按旧格式固定列
            if re.match(r'^\d+$', (cells.get('A') or '').strip()):
                cols = ('A', 'B', 'C', 'D')
            else:
                cols = _map_columns(cells) or ('A', 'B', 'C', 'D')
                continue

        key = clean_cell(cells.get(cols[0], ''))
        # 只接受正整数 key (object id)
        if not re.match(r'^\d+$', key):
            continue
        translations[key] = {
            'english': clean_cell(cells.get(cols[1], '')),
            'chinese': clean_cell(cells.get(cols[2], '')),
            'label':   normalize_suffix(clean_cell(cells.get(cols[3], ''))) if cols[3] else '',
        }
    return translations


def clean_cell(s):
    """单元格清洗: 去内部回车换行, 去前后空白"""
    if s is None:
        return ''
    s = str(s)
    s = s.replace('\r', '').replace('\n', '')
    return s.strip()


def normalize_suffix(s):
    """后缀规范: 每个 # 后必须有一个空格(行业规范, 提升码风)
    '#Flowering' -> '# Flowering';  '#W #cart' -> '# W # cart'
    行尾的 # 不处理(后面没内容)"""
    return re.sub(r'#(?=\S)', '# ', s or '')


# ---------- object 文件读写 ----------

def read_obj_line2(obj_path):
    """读 object 第2行(已去 \\r\\n); 文件不存在或不足2行返回 None"""
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
    """构建名称部分(不含后缀)"""
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
    """构建完整第2行: 名字 + 后缀 (后缀自带 #, 无空格)"""
    return build_name(mode, t) + t['label']


# ---------- 扫描 ----------

def scan(translations, mode, obj_dir):
    """扫描所有 key, 返回 (errors, warnings) —— 只检查名字, 不检查后缀"""
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


# ---------- 同步工具: 备份 / 撤回 ----------

BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'sync_backup')
BACKUP_XLSX = os.path.join(BACKUP_DIR, 'table_backup.xlsx')
BACKUP_OBJ = os.path.join(BACKUP_DIR, 'objects')
BACKUP_MARK = os.path.join(BACKUP_DIR, 'last_action.txt')


def save_table_backup():
    """备份翻译表, 标记最近一次动作为 table"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    shutil.copy2(DEFAULT_XLSX, BACKUP_XLSX)
    with open(BACKUP_MARK, 'w', encoding='utf-8') as f:
        f.write('table')


def save_objects_backup(keys):
    """备份将被改动的 object 文件, 标记最近一次动作为 objects"""
    os.makedirs(BACKUP_OBJ, exist_ok=True)
    for fn in os.listdir(BACKUP_OBJ):
        os.remove(os.path.join(BACKUP_OBJ, fn))
    for k in keys:
        shutil.copy2(os.path.join(DEFAULT_OBJ_DIR, f'{k}.txt'),
                     os.path.join(BACKUP_OBJ, f'{k}.txt'))
    with open(BACKUP_MARK, 'w', encoding='utf-8') as f:
        f.write('objects')


def action_undo():
    """撤回: 回退到上一次操作前的状态"""
    if not os.path.isfile(BACKUP_MARK):
        print('[撤回] 没有可撤回的操作')
        return
    with open(BACKUP_MARK, encoding='utf-8') as f:
        last = f.read().strip()
    if last == 'table':
        if not os.path.isfile(BACKUP_XLSX):
            print('[撤回] 表备份缺失, 无法撤回')
            return
        shutil.copy2(BACKUP_XLSX, DEFAULT_XLSX)
        os.remove(BACKUP_MARK)
        print('[撤回] 翻译表已回退到上一次操作前')
    elif last == 'objects':
        files = [f for f in os.listdir(BACKUP_OBJ) if f.endswith('.txt')]
        if not files:
            print('[撤回] 对象备份为空, 无法撤回')
            return
        for f in files:
            shutil.copy2(os.path.join(BACKUP_OBJ, f),
                         os.path.join(DEFAULT_OBJ_DIR, f))
        os.remove(BACKUP_MARK)
        print(f'[撤回] {len(files)} 个 object 已回退到上一次操作前')
    else:
        print(f'[撤回] 未知备份类型: {last}')


# ---------- 同步工具: 新增 / 导入修改 ----------

def split_cn_en(name):
    """'轨道终点Track End' -> ('轨道终点', 'Track End'); 纯中文/纯英文也能处理"""
    name = name.strip()
    i = 0
    while i < len(name) and ord(name[i]) > 127:
        i += 1
    return name[:i].rstrip(), name[i:].lstrip()


def parse_line2(line2):
    """'轨道终点Track End# W cart' -> (cn, en, suffix); 无#时后缀为空"""
    idx = line2.find('#')
    if idx < 0:
        name, suf = line2, ''
    else:
        name, suf = line2[:idx], line2[idx:]
    cn, en = split_cn_en(name)
    return cn, en, suf


def load_table():
    """加载翻译表(读写), 返回 (wb, ws, {key: 行号}); 表坏样式打不开时返回 None"""
    wb = guard_load_workbook()
    if wb is None:
        return None
    ws = wb[SHEET_NAME]
    key2row = {}
    for row in ws.iter_rows(min_col=1, max_col=1):
        k = str(row[0].value or '').strip()
        if k.isdigit():
            key2row[k] = row[0].row
    return wb, ws, key2row


def _find_status_col(ws):
    """按表头找修改状态列(1基); 找不到回退第8列(旧格式)"""
    for i in range(1, (ws.max_column or 9) + 1):
        v = ws.cell(row=1, column=i).value
        if v and '修改状态' in str(v):
            return i
    return 8


def action_add():
    """1. 新增物品: objects/ 里表上没有的编号, 追加为新行(B英文/C中文/D后缀)
    若日志确认某本地编号是ID转移目标, 则把表里旧号行 key 改成新号(保留翻译),
    而不是重复新增; 没有日志/日志无记录则照旧追加。"""
    loaded = load_table()
    if loaded is None:
        return
    wb, ws, key2row = loaded
    status_col = _find_status_col(ws)

    # ---- ID 转移: 日志确认的先改行 key, 再新增其余 ----
    translations = read_xlsx(DEFAULT_XLSX)
    det = detect_id_transfers(translations, DEFAULT_OBJ_DIR)
    plan = plan_table_transfers(det, translations)
    key2row, moved_dests, deleted, skipped = apply_plan(ws, key2row, plan)
    report_transfer_actions(det, moved_dests, deleted, skipped)

    added, skipped_add = [], []
    for fn in sorted(os.listdir(DEFAULT_OBJ_DIR)):
        if not fn.endswith('.txt'):
            continue
        key = fn[:-4]
        if not key.isdigit() or key in key2row:
            continue
        line2 = read_obj_line2(os.path.join(DEFAULT_OBJ_DIR, fn))
        if line2 is None or line2 == '':
            skipped_add.append((key, '第2行缺失或为空'))
            continue
        cn, en, suf = parse_line2(line2)
        if not cn and not en:
            skipped_add.append((key, '名字为空'))
            continue
        r = ws.max_row + 1
        ws.cell(row=r, column=1).value = int(key)
        ws.cell(row=r, column=2).value = en or None
        ws.cell(row=r, column=3).value = cn or None
        ws.cell(row=r, column=4).value = suf or None
        ws.cell(row=r, column=5).value = en or None
        ws.cell(row=r, column=6).value = cn or None
        ws.cell(row=r, column=7).value = suf or None
        ws.cell(row=r, column=status_col).value = '=IF(AND(B{r}=E{r},C{r}=F{r},D{r}=G{r}),"未修改","有修改")'.format(r=r)
        added.append(key)
    if added or moved_dests or deleted:
        save_table_backup()
        wb.save(DEFAULT_XLSX)
    print(f'[新增] 加入 {len(added)} 行: {", ".join(added[:20])}{" ..." if len(added) > 20 else ""}')
    if skipped_add:
        print(f'[新增] 跳过 {len(skipped_add)} 个: {skipped_add[:10]}{" ..." if len(skipped_add) > 10 else ""}')


def action_import():
    """2. 导入修改: 本地第2行与表(实装B/C/D)不一致的, 用本地覆盖 B/C/D
    若差异"像是ID转移"(本地X的内容=表里Y的内容)且 idTransferLog 确认 Y→…→X:
    改表行 key (Y→X, 保留B/C/D) 而不是覆盖名字; 孤儿行(物品已不在本地)删除并报告。
    没有日志/日志无对应记录 -> 算了, 按普通覆盖处理。"""
    loaded = load_table()
    if loaded is None:
        return
    wb, ws, key2row = loaded

    # ---- ID 转移: 日志确认的先改行 key, 不覆盖名字 ----
    translations = read_xlsx(DEFAULT_XLSX)
    det = detect_id_transfers(translations, DEFAULT_OBJ_DIR)
    plan = plan_table_transfers(det, translations)
    key2row, moved_dests, deleted, skipped = apply_plan(ws, key2row, plan)
    report_transfer_actions(det, moved_dests, deleted, skipped)

    changed = []
    for k in sorted(key2row, key=int):
        if k in moved_dests:
            continue  # 转移目标行: 保留原翻译, 不覆盖
        r = key2row[k]
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        d = ws.cell(row=r, column=4).value
        t = {
            'english': str(b).strip() if b is not None else '',
            'chinese': str(c).strip() if c is not None else '',
            'label': normalize_suffix(str(d).strip() if d is not None else ''),
        }
        line2 = read_obj_line2(os.path.join(DEFAULT_OBJ_DIR, f'{k}.txt'))
        if line2 is None:
            continue  # object 不存在, 不管
        if build_line2(3, t) == line2:
            continue  # 表与本地一致
        cn, en, suf = parse_line2(line2)
        ws.cell(row=r, column=2).value = en or None
        ws.cell(row=r, column=3).value = cn or None
        ws.cell(row=r, column=4).value = suf or None
        changed.append(k)
    if changed or moved_dests or deleted:
        save_table_backup()
        wb.save(DEFAULT_XLSX)
    print(f'[导入修改] 用本地覆盖 {len(changed)} 行: {", ".join(changed[:20])}{" ..." if len(changed) > 20 else ""}')


def action_delete_old():
    """9. 删除旧翻译: 翻译表里有、本地 objects 没有的编号 -> 删行 (备份后可撤回)"""
    loaded = load_table()
    if loaded is None:
        return
    wb, ws, key2row = loaded

    local_ids = set()
    for fn in os.listdir(DEFAULT_OBJ_DIR):
        if fn.endswith('.txt') and fn[:-4].isdigit():
            local_ids.add(fn[:-4])

    old = sorted((k for k in key2row if k not in local_ids), key=int)
    if not old:
        print('没有可删除的旧翻译: 翻译表里所有编号本地都有对应 objects。')
        return

    print(f'翻译表 {len(key2row)} 行 | 本地 objects {len(local_ids)} 个')
    print(f'本地没有、表中多余的旧编号 {len(old)} 个: '
          f'{", ".join(old[:20])}{" ..." if len(old) > 20 else ""}')
    try:
        ans = input('确认删除这些旧翻译行? (y 删除 / 其他键取消): ').strip().lower()
    except EOFError:
        print()
        return
    if ans != 'y':
        print('已取消, 没有改动。')
        return
    for r in sorted((key2row[k] for k in old), reverse=True):
        ws.delete_rows(r, 1)
    save_table_backup()
    wb.save(DEFAULT_XLSX)
    print(f'[删除旧翻译] 已删除 {len(old)} 行。')


# ---------- 执行翻译 ----------

def run_mode(mode, ignore_errors=False, backup_objects=False, dry_run=False):
    """执行翻译(1/2/3); backup_objects=True 时先把将改动的 object 备份(供撤回)"""
    xlsx_path = DEFAULT_XLSX
    obj_dir = DEFAULT_OBJ_DIR

    print(f"模式: {mode_label(mode)}")
    print(f"读取 xlsx: {xlsx_path}")
    if not os.path.isfile(xlsx_path):
        print(f"错误: 找不到 xlsx 文件 '{xlsx_path}'")
        return

    print("解析 xlsx ...")
    translations = read_xlsx(xlsx_path)
    print(f"xlsx 共 {len(translations)} 条 key")

    # ---- ID 转移警告: 表与本地编号未对齐时, 应用翻译会还原转移后的名字 ----
    det = detect_id_transfers(translations, obj_dir)
    if det['confirmed']:
        print('\n[警告] 检测到日志确认的ID转移, 表与本地编号尚未对齐:')
        for y, x, path in det['confirmed']:
            chain = '→'.join(str(a) for a, b in path) + f'→{x}'
            print(f'    {y}->{x} ({chain}): 物品已改号, 表里还在旧号')
        print('    继续覆盖会把转移后物品的名字按旧表还原。')
        print('    建议先执行菜单 2. 导入修改 同步表, 再回来应用翻译。')

    print("扫描 object 第2行 ...")
    errors, warnings = scan(translations, mode, obj_dir)

    if warnings:
        print(f"\n--- warnings: {len(warnings)} 条 ---")
        for w in warnings:
            print(f"  [WARN] {w}")

    if errors:
        print(f"\n--- 异常: {len(errors)} 条 ---")
        for e in errors:
            print(f"  [ERROR] {e}")
        if ignore_errors:
            print(f"\n--ignore-errors: 忽略 {len(errors)} 处异常, 继续覆盖 "
                  f"(将用 xlsx 数据覆盖 object, 可修复空行)。")
        else:
            print(f"\n扫描发现 {len(errors)} 处异常, 终止覆盖。"
                  f"请先修复, 或加 --ignore-errors 忽略异常继续覆盖。")
            return

    print(f"\n扫描通过: {len(translations)} 条, warning {len(warnings)} 条, "
          f"异常 {len(errors)} 条"
          f"{' (已忽略)' if ignore_errors else ''}。")

    changed = []
    for key in sorted(translations.keys(), key=lambda x: int(x)):
        t = translations[key]
        if not has_name(mode, t):
            continue
        line2 = read_obj_line2(os.path.join(obj_dir, f'{key}.txt'))
        if line2 is None or line2 == '':
            continue
        if line2 != build_line2(mode, t):
            changed.append(key)
    print(f"\n预览: 覆盖后将改动 {len(changed)} 个 object 的第2行"
          f"{' (--dry-run, 未写入)' if dry_run else ''}")
    for key in changed[:10]:
        t = translations[key]
        line2 = read_obj_line2(os.path.join(obj_dir, f'{key}.txt'))
        print(f"  {key}: {line2!r} -> {build_line2(mode, t)!r}")

    if dry_run:
        print("--dry-run 模式, 不执行覆盖。")
        return

    if backup_objects and changed:
        save_objects_backup(changed)

    print("开始覆盖 ...")
    count, skipped = translate(translations, mode, obj_dir)
    print(f"完成: 覆盖 {count} 个 object, 跳过 {skipped} 个。")


# ---------- 云同步 (腾讯文档共享表) ----------

QQSYNC_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'claude AI', 'qqdoc-sync')


def _qqsync():
    """懒加载 claude AI/qqdoc-sync/qqdoc_sync.py (配置/token/备份都留在那边, 不进本仓库)。"""
    if QQSYNC_DIR not in sys.path:
        sys.path.insert(0, QQSYNC_DIR)
    try:
        import qqdoc_sync
    except ImportError as e:
        print(f'云同步模块加载失败: {e}')
        print(f'(期望目录: {QQSYNC_DIR})')
        return None
    return qqdoc_sync


def _run_qq(pick):
    """跑一个云同步入口; 出错转成人话, 不炸回菜单。"""
    qs = _qqsync()
    if qs is None:
        return
    try:
        pick(qs)
    except SystemExit as e:  # 引擎 die() 走 sys.exit
        if e.code not in (0, None):
            print(f'云同步中止 (exit {e.code}), 详见上方报错。')
    except RuntimeError as e:
        print(f'云同步出错: {e}')
    except KeyboardInterrupt:
        print('\n已中断。')


# ---------- ID 转移检测 (idTransferLog.txt) ----------

def find_id_transfer_logs():
    """候选 idTransferLog.txt 位置列表, 按优先级: 环境变量 ID_TRANSFER_LOG
    (指定后只含它) > 当前运行目录 > OneLifeData7 脚本目录 > 仓库根 >
    仓库根/output。全部合并读取, 日志搬家后新旧双份并存也能读全。"""
    env = os.environ.get('ID_TRANSFER_LOG', '').strip()
    if env:
        return [env] if os.path.isfile(env) else []
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(script_dir)
    paths = []
    for d in [os.getcwd(), script_dir, root, os.path.join(root, 'output')]:
        try:
            for fn in os.listdir(d):
                if fn.lower() == 'idtransferlog.txt':
                    paths.append(os.path.join(d, fn))
        except OSError:
            continue
    return paths


def find_id_transfer_log():
    """返回优先级最高的日志路径; 一个都找不到返回 None"""
    paths = find_id_transfer_logs()
    return paths[0] if paths else None


def _expand_range(s):
    """'123~127' -> [123..127]; '123' -> [123]; 解析不了返回 []"""
    s = s.strip()
    m = re.match(r'^(\d+)\s*~\s*(\d+)$', s)
    if m:
        a, b = sorted((int(m.group(1)), int(m.group(2))))
        return list(range(a, b + 1))
    if re.match(r'^\d+$', s):
        return [int(s)]
    return []


def load_id_transfer_log(path=None):
    """解析 idTransferLog.txt -> {旧ID: 新ID}。行格式 'A→B' 或
    '123~127→369~374'(号段按位配对), 兼容 '->'。path 未给时合并读取全部
    候选位置(同一旧ID以优先级高的文件为准); 都找不到返回 {}。"""
    if path is None:
        merged = {}
        for p in find_id_transfer_logs():
            for k, v in load_id_transfer_log(p).items():
                merged.setdefault(k, v)
        return merged
    if not os.path.isfile(path):
        return {}
    mapping = {}
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            sep = '→' if '→' in line else ('->' if '->' in line else None)
            if sep is None:
                continue
            left, _, right = line.partition(sep)
            olds, news = _expand_range(left), _expand_range(right)
            if not olds or not news:
                continue
            for i, old in enumerate(olds):
                mapping[old] = news[i] if i < len(news) else news[-1]
    return mapping


def follow_transfer(mapping, old):
    """沿转移链走到最终编号: 返回 (路径[(a,b)...], 终点); 无记录/成环返回 (None, None)"""
    path, cur, seen = [], old, set()
    while cur in mapping:
        if cur in seen:
            return None, None
        seen.add(cur)
        nxt = mapping[cur]
        path.append((cur, nxt))
        cur = nxt
    return (path, cur) if path else (None, None)


def detect_id_transfers(translations, obj_dir):
    """比对本地 objects 与翻译表: 只有"两边都反过来"才算疑似反转 ——
    本地 X 与 Y 的名字在表里互相交换 (本地X=表Y 且 本地Y=表X, X≠Y),
    单边撞名不算。再用 idTransferLog 确认 (日志里两个编号之间存在改号记录)。

    返回 dict:
      log_path:   日志文件路径, 找不到为 None
      log:        {旧ID: 新ID}, 无日志为 {}
      local_line2: {X: 本地第2行}
      candidates:  [(X, Y, line2)] 疑似反转对 (X<Y)
      confirmed:   [(旧ID, 新ID, path)]  日志有对应记录 = 物品ID转移
      unconfirmed: [(X, Y)]        日志无对应记录
      name_same:  本地与表名字相同(仅后缀差异)而跳过的编号数
    """
    log_path = find_id_transfer_log()
    log = load_id_transfer_log(log_path)

    local_line2 = {}
    for fn in os.listdir(obj_dir):
        if not fn.endswith('.txt') or not fn[:-4].isdigit():
            continue
        line2 = read_obj_line2(os.path.join(obj_dir, fn))
        if line2:
            local_line2[fn[:-4]] = line2

    by_name = {}
    for k, t in translations.items():
        l2 = build_line2(3, t)
        if l2:
            by_name.setdefault(l2, []).append(k)

    def table_line2(k):
        t = translations.get(k)
        return build_line2(3, t) if t else None

    EMPTY_T = {'english': '', 'chinese': '', 'label': ''}
    candidates, confirmed, unconfirmed = [], [], []
    name_same = 0
    for x, l2 in sorted(local_line2.items(), key=lambda kv: int(kv[0])):
        if l2 == table_line2(x):
            continue  # 本地与表完全一致, 不疑似
        t_x = translations.get(x)
        if t_x:
            # 两边名字相同(不含后缀) -> 只是后缀差异, 不是转移, 不用管
            cn, en, suf = parse_line2(l2)
            if cn + en == build_name(3, t_x):
                name_same += 1
                continue
        txl2 = table_line2(x)
        if txl2 is None:
            continue  # 表里没有 X 行, 无从反转
        xi = int(x)
        for y in by_name.get(l2, []):
            yi = int(y)
            if yi <= xi:
                continue  # 每对只记一次 (X<Y)
            # 反转 = 两边都反过来: 本地 X 的名字在表 Y 上, 且本地 Y 的名字在表 X 上
            if txl2 != local_line2.get(y):
                continue  # 单边撞名不算反转
            candidates.append((x, y, l2))
            path, final = follow_transfer(log, xi)
            if final == yi:
                confirmed.append((x, y, path))
                continue
            path, final = follow_transfer(log, yi)
            if final == xi:
                confirmed.append((y, x, path))
                continue
            unconfirmed.append((x, y))
    return {'log_path': log_path, 'log': log, 'local_line2': local_line2,
            'candidates': candidates, 'confirmed': confirmed,
            'unconfirmed': unconfirmed, 'name_same': name_same}


def plan_table_transfers(det, translations):
    """基于检测结果规划表修复动作 (不改表):
      ('move', Y, X)          源行 Y 的 key 改为 X (B/C/D 不动)
      ('delete_orphan', K)    孤儿行 K: 其物品已不在本地, 删行
      ('keep_orphan', K)      疑似孤儿但内容仍存在于本地某处, 保留待人工
      ('skip', Y, X, 原因)    日志确认但无法自动处理
    返回动作列表 (move 按依赖排序, 先删后移)。"""
    local_contents = set(det['local_line2'].values())
    srcs = {int(y) for y, x, p in det['confirmed']}
    dsts = {int(x) for y, x, p in det['confirmed']}
    handled_orphan = set()
    actions = []

    def orphan_act(k):
        k = str(k)
        if k in handled_orphan:
            return []
        handled_orphan.add(k)
        if k not in translations:
            return []
        content = build_line2(3, translations[k])
        if content and content in local_contents:
            return [('keep_orphan', k)]
        return [('delete_orphan', k)]

    for y, x, path in det['confirmed']:
        if y not in translations:
            actions.append(('skip', y, x, '源行不在翻译表中'))
            continue
        if x in translations:
            content = build_line2(3, translations[x])
            if content and content in local_contents:
                # 目标已有行且其内容仍存在于本地: 两边都按普通差异覆盖即可
                continue
            actions.extend(orphan_act(x))
        actions.append(('move', y, x))

    # 依赖排序: 某 move 的目标是另一 move 的源时, 先移掉占位的那个
    moves = [a for a in actions if a[0] == 'move']
    by_dst = {a[2]: a for a in moves}
    state, ordered = {}, []

    def visit(m):
        s = state.get(m[1])
        if s == 2:
            return True
        if s == 1:
            return False
        state[m[1]] = 1
        dep = by_dst.get(m[2])
        if dep is not None and dep != m and not visit(dep):
            return False
        state[m[1]] = 2
        ordered.append(m)
        return True

    for m in moves:
        if not visit(m):
            for m2 in [a for a in moves if a not in ordered]:
                actions.remove(m2)
                actions.append(('skip', m2[1], m2[2], '转移链成环, 请人工处理'))
            break

    return [a for a in actions if a[0] != 'move'] + ordered


def rebuild_key2row(ws):
    """重建 {key: 行号} (删行后行号会变, 用这个刷新)"""
    key2row = {}
    for row in ws.iter_rows(min_col=1, max_col=1):
        k = str(row[0].value or '').strip()
        if k.isdigit():
            key2row[k] = row[0].row
    return key2row


def apply_plan(ws, key2row, plan):
    """执行计划: 先按行号从大到小删孤儿行并重建 key2row, 再改 A 列 key。
    返回 (新 key2row, 转移目标集合, 删除数, 跳过列表)。"""
    del_rows = sorted((key2row[a[1]] for a in plan
                       if a[0] == 'delete_orphan' and a[1] in key2row),
                      reverse=True)
    for r in del_rows:
        ws.delete_rows(r, 1)
    if del_rows:
        key2row = rebuild_key2row(ws)
    moved_dests = set()
    for a in plan:
        if a[0] == 'move' and a[1] in key2row:
            ws.cell(row=key2row[a[1]], column=1).value = int(a[2])
            moved_dests.add(a[2])
    if moved_dests:
        key2row = rebuild_key2row(ws)
    skipped = [a for a in plan if a[0] == 'skip']
    return key2row, moved_dests, len(del_rows), skipped


def report_transfer_actions(det, moved_dests, deleted, skipped):
    """打印 ID 转移处理摘要 (菜单1/2 执行后调用)"""
    if det['log_path'] is None and not det['candidates']:
        return
    if det['log_path'] is None:
        print(f'[ID转移] 未找到 idTransferLog.txt, 无法确认'
              f' {len(det["candidates"])} 处疑似反转')
        return
    print(f'[ID转移] 日志 {det["log_path"]} ({len(det["log"])} 条) | '
          f'疑似 {len(det["candidates"])} | 确认 {len(det["confirmed"])} | '
          f'未确认 {len(det["unconfirmed"])}')
    for y, x, path in det['confirmed']:
        if x in moved_dests:
            print(f'   确认 {y} 与 {x} 疑似反转, 日志有对应记录: 表行 key 已改为 {x}, 保留B/C/D翻译')
        else:
            print(f'   确认 {y} 与 {x} 疑似反转, 日志有对应记录')
    for x, y in det['unconfirmed']:
        print(f'   未确认 {x} 与 {y} 疑似反转, 日志无对应记录')
    if deleted:
        print(f'   已删除 {deleted} 个孤儿行 (其物品已不在本地)')
    for a in skipped:
        print(f'   跳过 {a[1]} 与 {a[2]}: {a[3]}')


def action_check_transfers():
    """只读: 检测本地与表的疑似反转编号, 用 idTransferLog 确认ID转移, 并预览修复动作"""
    xlsx_path = DEFAULT_XLSX
    if not os.path.isfile(xlsx_path):
        print(f'错误: 找不到 xlsx 文件 {xlsx_path}')
        return
    translations = read_xlsx(xlsx_path)
    det = detect_id_transfers(translations, DEFAULT_OBJ_DIR)
    print('===== ID 转移检测 (只读, 不写任何文件) =====')
    print(f'翻译表: {len(translations)} 行 | 本地 objects: {len(det["local_line2"])} 个')
    if det['log_path']:
        print(f'转移日志: {det["log_path"]} ({len(det["log"])} 条记录)')
    else:
        print('转移日志: 未找到 idTransferLog.txt (已查 当前目录/OneLifeData7/仓库根/output)')
    if not det['candidates']:
        print('未发现"两边名字反过来"的疑似反转。')
        if det['log']:
            print(f'(日志 {len(det["log"])} 条记录当前无对应反转迹象, 可能已同步过。)')
        return
    print(f'--- 疑似反转: {len(det["candidates"])} 处 ---')
    for x, y, l2 in det['candidates']:
        print(f'  [疑似] {x} 与 {y} 疑似反转')
    if det['name_same']:
        print(f'(本地与表名字相同、仅后缀差异的 {det["name_same"]} 处不比对, 已跳过)')
    print('--- 日志确认 ---')
    for y, x, path in det['confirmed']:
        print(f'  [ID转移] {y} 与 {x} 疑似反转, 日志有对应记录')
    for x, y in det['unconfirmed']:
        print(f'  [未确认] {x} 与 {y} 疑似反转, 日志无对应记录')
    if det['log_path'] is None:
        print('(没有日志, 全部疑似反转都按普通差异处理)')
    plan = plan_table_transfers(det, translations)
    if plan:
        print('--- 修复预览 (在菜单 1/2 执行时应用) ---')
        for a in plan:
            if a[0] == 'move':
                print(f'  [改key] 表行 {a[1]} 改为 {a[2]} (保留B/C/D翻译)')
            elif a[0] == 'delete_orphan':
                print(f'  [删孤儿行] 表行 {a[1]} (其物品已不在本地)')
            elif a[0] == 'keep_orphan':
                print(f'  [保留] 表行 {a[1]} (内容仍存在于本地某处, 请人工核对)')
            else:
                print(f'  [跳过] {a[1]} 与 {a[2]}: {a[3]}')


# ---------- 同步工具菜单 ----------

def sync_menu():
    """无参数启动时的菜单: 新增物品/导入修改/应用翻译/撤回"""
    wb = guard_load_workbook(read_only=True)
    if wb is not None:
        ws = wb[SHEET_NAME]
        n_obj = len([f for f in os.listdir(DEFAULT_OBJ_DIR)
                     if f.endswith('.txt') and f[:-4].isdigit()])
        print(f'翻译表: {ws.max_row - 1} 行 | 本地 objects: {n_obj} 个')
        wb.close()

    while True:
        print()
        print('======== 翻译表同步工具 ========')
        print('1. 新增物品   (本地objects -> 本地翻译表)')
        print('2. 导入修改   (本地objects -> 本地翻译表)')
        print('3. 应用翻译   (本地翻译表 -> 本地objects, 可选翻译模式)')
        print('4. 撤回       (回退到上一次操作前)')
        print('-------- 云同步 (腾讯文档共享表) --------')
        print('5. 云同步对比 (对比远程, 只读)')
        print('6. 拉取翻译表 (共享翻译表 -> 本地翻译表)')
        print('7. 推送翻译表 (本地翻译表 -> 共享翻译表)')
        print('-------- 其他 --------')
        print('8. 转移检测   (只读: 比对本地与表疑似反转的编号, 用 idTransferLog 确认)')
        print('9. 删除旧翻译 (删除本地已没有物品的表行, 备份后可撤回)')
        print('0. 退出')
        try:
            choice = input('请选择: ').strip()
        except EOFError:
            print()
            break
        if choice == '0':
            break
        elif choice == '1':
            action_add()
        elif choice == '2':
            action_import()
        elif choice == '3':
            try:
                m = input('翻译模式 (1 翻译英文 / 2 翻译中文 / 3 追加英文, 回车默认3): ').strip()
            except EOFError:
                print()
                return
            if m == '' or m == '3':
                run_mode(3, backup_objects=True)
            elif m in ('1', '2'):
                run_mode(int(m), backup_objects=True)
            else:
                print('无效选择')
        elif choice == '4':
            action_undo()
        elif choice == '5':
            _run_qq(lambda qs: qs.menu_status())
        elif choice == '6':
            _run_qq(lambda qs: qs.menu_pull())
        elif choice == '7':
            _run_qq(lambda qs: qs.menu_push())
        elif choice == '8':
            action_check_transfers()
        elif choice == '9':
            action_delete_old()
        else:
            print('无效选择')


# ---------- main ----------

def main():
    if '--check-transfer' in sys.argv:
        action_check_transfers()
        return
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
        sync_menu()
        return
    run_mode(mode, ignore_errors=ignore_errors, dry_run=dry_run)


if __name__ == '__main__':
    main()

"""
批量银行报盘合并与工资表打印工具
"""

import os
import shutil
import tkinter as tk
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from tkinter import filedialog, messagebox, scrolledtext

import openpyxl
import xlrd

HEADERS = ["序号", "账户", "户名", "金额", "跨行标识", "行名", "联行行号", "摘要", "备注"]


def split_filename(filename):
    """
    解析文件名: YYYYMM-银行名称-单位名.xls
    返回 (yearmon, bank_name, unit_name)
    注意: 单位名可能含 '-'，如 (吉林大学）深部探测与成像全国重点实验室
    策略: 第一个 '-' 前是年月，第二个 '-' 前是银行名，之后全是单位名
    """
    name_no_ext = filename.rsplit(".", 1)[0]
    parts = name_no_ext.split("-", 2)
    if len(parts) != 3:
        raise ValueError(f"文件名格式不匹配: {filename}")
    return parts[0], parts[1], parts[2]


def detect_bank_type(filename):
    """
    从文件名判断银行类型
    返回: 'ccb' (建设银行), 'icbc' (工商银行), 'jlb' (吉林银行)
    """
    if "建设银行" in filename:
        return "ccb"
    if "工商银行" in filename:
        return "icbc"
    if "吉林银行" in filename:
        return "jlb"
    raise ValueError(f"无法识别银行类型: {filename}")


def rename_bank_files(bank_dir, output_dir):
    """
    扫描 bank_dir 下所有 .xls 文件
    改名规则: 单位名-YYYYMM-银行名称.xls
    复制到 output_dir（不修改源文件）
    返回: [(new_name, old_name, bank_name, unit_name), ...]
    """
    os.makedirs(output_dir, exist_ok=True)
    result = []

    for fname in os.listdir(bank_dir):
        if not fname.lower().endswith(".xls"):
            continue

        old_path = os.path.join(bank_dir, fname)
        yearmon, bank_name, unit_name = split_filename(fname)

        # 新文件名: 单位名-YYYYMM-银行名称.xls
        new_name = f"{unit_name}-{yearmon}-{bank_name}.xls"
        new_path = os.path.join(output_dir, new_name)

        shutil.copy2(old_path, new_path)
        result.append((new_name, fname, bank_name, unit_name))

    return result


def _to_decimal(val):
    """将 xlrd 读取的值转为 Decimal（两位小数，分以下直接舍去），避免 float 精度问题"""
    try:
        if isinstance(val, str):
            val = val.replace(",", "").strip()
        d = Decimal(str(val))
        # 检查是否有分以下的数值（厘）
        truncated = d.quantize(Decimal("0.01"), rounding=ROUND_DOWN)
        if truncated != d:
            return truncated, True  # (值, 有警告)
        return truncated, False
    except Exception:
        return Decimal("0.00"), False


def _read_icbc_rows(filepath, warnings):
    """读取工商银行 3 列格式，映射为建设银行 9 列"""
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_index(0)
    rows = []
    for r in range(1, ws.nrows):  # 跳过表头
        seq = len(rows) + 1
        account = str(ws.cell_value(r, 0)).strip()
        name = str(ws.cell_value(r, 1)).strip()
        amount, warned = _to_decimal(ws.cell_value(r, 2))
        if warned:
            warnings.append(f"{os.path.basename(filepath)} 第{r+1}行 {name} 金额 {ws.cell_value(r, 2)} 含分以下数值，已舍去")
        rows.append([seq, account, name, amount, "1", "工商银行", "", "", ""])
    return rows


def _read_ccb_rows(filepath, warnings):
    """读取建设银行 9 列格式，按模板要求：跨行标识填0，行名留空"""
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_index(0)
    rows = []
    for r in range(1, ws.nrows):  # 跳过表头
        seq = len(rows) + 1
        row = []
        for c in range(9):
            val = ws.cell_value(r, c)
            # 金额列（索引3）转数字
            if c == 3:
                val, warned = _to_decimal(val)
                if warned:
                    name = str(ws.cell_value(r, 2)).strip()
                    warnings.append(f"{os.path.basename(filepath)} 第{r+1}行 {name} 金额 {ws.cell_value(r, 3)} 含分以下数值，已舍去")
            row.append(val)
        # 序号重新生成
        row[0] = seq
        # 模板要求：建行跨行标识填0，行名留空
        row[4] = "0"
        row[5] = ""
        rows.append(row)
    return rows


def _read_jlb_rows(filepath, warnings):
    """读取吉林银行 9 列格式，按模板要求：跨行标识填1，行名有值用值无则'吉林银行'"""
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_index(0)
    rows = []
    for r in range(1, ws.nrows):  # 跳过表头
        seq = len(rows) + 1
        row = []
        for c in range(9):
            val = ws.cell_value(r, c)
            # 金额列（索引3）转数字
            if c == 3:
                val, warned = _to_decimal(val)
                if warned:
                    name = str(ws.cell_value(r, 2)).strip()
                    warnings.append(f"{os.path.basename(filepath)} 第{r+1}行 {name} 金额 {ws.cell_value(r, 3)} 含分以下数值，已舍去")
            row.append(val)
        # 序号重新生成
        row[0] = seq
        # 模板要求：他行跨行标识填1
        row[4] = "1"
        # 行名：有值用值，无则"吉林银行"
        bank_name = str(row[5]).strip()
        if not bank_name:
            row[5] = "吉林银行"
        rows.append(row)
    return rows


def _bank_sort_key(bank_name):
    """银行排序：建设银行 < 工商银行 < 吉林银行"""
    order = {"建设银行": 0, "工商银行": 1, "吉林银行": 2}
    return order.get(bank_name, 99)


def merge_bank_files(renamed_list, bank_dir, output_dir):
    """
    按单位名分组，每个单位一个合并文件
    建设银行/吉林银行: 直接复制9列
    工商银行: 3列 → 映射到9列
    合并文件名: {单位名}-{年月}-{银行1}{行数}-{银行2}{行数}.xlsx
    银行按字母排序（建设银行 < 工商银行 < 吉林银行）
    返回: (merged_files, warnings) 已排序
    """
    os.makedirs(output_dir, exist_ok=True)

    # 按单位名分组
    units = {}
    for new_name, old_name, bank_name, unit_name in renamed_list:
        units.setdefault(unit_name, []).append((new_name, old_name, bank_name))

    merged_files = []
    warnings = []

    # 计算序号位数（根据文件总数）
    total_units = len(units)
    pad_width = len(str(total_units))

    for idx, (unit_name, entries) in enumerate(sorted(units.items()), start=1):
        # 按银行分组
        bank_groups = {}
        yearmon = None
        for new_name, old_name, bank_name in entries:
            bank_groups.setdefault(bank_name, []).append(old_name)
            # 从第一个 entry 获取年份
            if yearmon is None:
                y, _, _ = split_filename(old_name)
                yearmon = y

        # 读取每个银行的数据
        all_rows = []
        bank_summary = []
        for bank_name in sorted(bank_groups, key=_bank_sort_key):
            file_list = bank_groups[bank_name]
            bank_rows = []
            for old_name in file_list:
                filepath = os.path.join(bank_dir, old_name)
                bt = detect_bank_type(old_name)
                if bt == "icbc":
                    bank_rows.extend(_read_icbc_rows(filepath, warnings))
                elif bt == "ccb":
                    bank_rows.extend(_read_ccb_rows(filepath, warnings))
                elif bt == "jlb":
                    bank_rows.extend(_read_jlb_rows(filepath, warnings))
            # 重新编号（跨银行连续递增）
            for i, row in enumerate(bank_rows):
                row[0] = len(all_rows) + i + 1
            all_rows.extend(bank_rows)
            bank_summary.append((bank_name, len(bank_rows)))

        # 构建合并文件名（加序号前缀 + 总计金额后缀）
        seq_prefix = str(idx).zfill(pad_width)
        parts = [f"{seq_prefix}_{unit_name}", yearmon]
        for bank_name, count in bank_summary:
            parts.append(f"{bank_name}{count}个")

        # 计算总计金额（元角分，Decimal 精确计算）
        total_fen = 0
        for row in all_rows:
            amt = row[3]
            if isinstance(amt, Decimal):
                total_fen += int(amt * 100)
            else:
                total_fen += int(round(float(amt) * 100))
        yuan = total_fen // 100
        jiao = (total_fen % 100) // 10
        fen = total_fen % 10

        merged_name = "-".join(parts) + f"-总计{yuan}元{jiao}角{fen}分.xlsx"
        merged_path = os.path.join(output_dir, merged_name)

        # 写入
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = unit_name
        ws.append(HEADERS)
        for row in all_rows:
            ws.append(row)
        wb.save(merged_path)
        merged_files.append(merged_name)

    return sorted(merged_files), warnings


def match_payroll_files(merged_files_list, payroll_dir):
    """
    输入：合并后的银行报盘文件列表（已排序）和工资表目录
    返回：[(merged_filename, payroll_filepath, unit_name), ...]
    与 merged_files_list 一一对应

    匹配逻辑：
    - 从合并文件名提取单位名（第一个 '-202606-' 之前的部分）
    - signed_ 版本优先
    - 排除 汇总表 和 验证 相关文件
    - 支持 .xls 格式
    """
    result = []

    # 扫描工资表目录，建立索引
    payroll_files = []
    for fname in os.listdir(payroll_dir):
        # 排除汇总表和验证文件
        if "汇总表" in fname or "验证" in fname:
            continue
        payroll_files.append(fname)

    # 按单位名建立索引：unit_name -> [(is_signed, filename), ...]
    # 先按 signed_ 分组，再按扩展名优先级排序
    from collections import defaultdict
    payroll_index = defaultdict(list)

    for fname in payroll_files:
        unit_candidate = None
        is_signed = fname.startswith("signed_")

        # 去掉 signed_ 前缀用于匹配
        base = fname[len("signed_"):] if is_signed else fname

        # 从工资表文件名中提取单位名（去掉年月后缀部分）
        # 格式: 吉林大学数学学院2026年06月工资表.xlsx → 吉林大学数学学院
        # 或: (吉林大学）深部探测与成像全国重点实验室202606工资.xls → (吉林大学）深部探测与成像全国重点实验室
        for sep in ["2026年06月工资表", "202606工资表", "202606工资"]:
            if sep in base:
                unit_candidate = base.split(sep, 1)[0].strip()
                break

        if unit_candidate:
            payroll_index[unit_candidate].append((is_signed, fname))

    # 对每个合并文件进行匹配
    matched_units = set()
    duplicates = {}  # unit_name -> [all matched filenames]
    for merged_name in merged_files_list:
        # 提取单位名：第一个 '-202606-' 之前的部分（去掉序号前缀）
        raw_unit = merged_name.split("-202606-", 1)[0]
        # 去掉开头的序号前缀 "001_"
        unit_name = raw_unit.split("_", 1)[1] if "_" in raw_unit else raw_unit

        matched_file = None

        if unit_name in payroll_index:
            matched_units.add(unit_name)
            candidates = payroll_index[unit_name]
            # 记录所有候选文件，用于重复检测
            all_candidates = [f for _, f in candidates]
            if len(all_candidates) > 1:
                duplicates[unit_name] = all_candidates

            # 优先级：signed_ > 非 signed_，同优先级下 .xlsx > .xls
            signed = [f for s, f in candidates if s]
            unsigned = [f for s, f in candidates if not s]

            if signed:
                # signed_ 中优先选 .xlsx
                xlsx = [f for f in signed if f.endswith(".xlsx")]
                matched_file = xlsx[0] if xlsx else signed[0]
            elif unsigned:
                xlsx = [f for f in unsigned if f.endswith(".xlsx")]
                matched_file = xlsx[0] if xlsx else unsigned[0]

        filepath = os.path.join(payroll_dir, matched_file) if matched_file else None
        result.append((merged_name, filepath, unit_name))

    # 反向检查：工资表中有哪些单位没被匹配到
    unmatched = []
    for unit_name, candidates in payroll_index.items():
        if unit_name not in matched_units:
            # 取第一个文件名作为代表
            _, fname = candidates[0]
            unmatched.append(fname)

    return result, unmatched, duplicates


# ──────────────────────────────────────────────
# WPS 打印模块
# ──────────────────────────────────────────────


def check_wps_available():
    """
    检测 WPS Office COM 组件是否可用
    返回: True/False
    """
    try:
        import win32com.client
        app = win32com.client.Dispatch("KET.Application")
        app.Quit()
        return True
    except (ImportError, Exception):
        return False


def print_file(filepath, progress_callback=None):
    """
    打印单个 Excel 文件
    参数:
      filepath: 文件路径
      progress_callback: 可选回调函数，用于 GUI 进度更新
    返回: True/False
    """
    import time

    try:
        import win32com.client
    except ImportError:
        if progress_callback:
            progress_callback(f"win32com 不可用（非 Windows 环境）")
        return False

    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            app = win32com.client.Dispatch("KET.Application")
            app.Visible = False

            wb = app.Workbooks.Open(filepath)
            ws = wb.ActiveSheet

            # 页面设置：A4 横向
            ws.PageSetup.Orientation = 2  # xlLandscape
            ws.PageSetup.PaperSize = 9    # xlPaperA4

            ws.PrintOut()

            wb.Close(SaveChanges=False)
            # 不退出 WPS 应用，保留以备多次打印

            return True

        except Exception as e:
            if progress_callback:
                progress_callback(f"打印失败 (第{attempt}次): {e}")
            if attempt < max_retries:
                time.sleep(5)
            else:
                return False
        finally:
            time.sleep(2.5)


def batch_print(matched_pairs, progress_callback=None):
    """
    批量打印工资表
    参数:
      matched_pairs: [(merged_filename, payroll_filepath, unit_name), ...]
      progress_callback: 可选回调函数(current, total, message)
    返回: (success_count, fail_count, fail_list)
    """
    if not check_wps_available():
        return (0, 0, ["WPS不可用"])

    total = len(matched_pairs)
    success_count = 0
    fail_count = 0
    fail_list = []

    for i, (merged_filename, payroll_filepath, unit_name) in enumerate(matched_pairs):
        if progress_callback:
            progress_callback(i + 1, total, f"正在打印: {unit_name}")

        ok = print_file(payroll_filepath)
        if ok:
            success_count += 1
        else:
            fail_count += 1
            fail_list.append((merged_filename, unit_name))

    return (success_count, fail_count, fail_list)


def generate_report_xlsx(output_dir, renamed, matched, unmatched, duplicates,
                         merge_warnings=None, success=None, fail=None, fail_list=None):
    """
    生成合并打印操作记录 xlsx
    renamed: [(new_name, old_name, bank_name, unit_name), ...]
    matched: [(merged_filename, payroll_filepath, unit_name), ...]
    unmatched: [payroll_filename, ...]
    duplicates: {unit_name: [filenames, ...], ...}
    merge_warnings: [str, ...] 金额警告
    success/fail/fail_list: 打印结果（仅打印模式有值）
    返回: 报告文件路径
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(output_dir, f"操作记录_{ts}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "操作记录"

    # ── 标题行 ──
    ws.cell(row=1, column=1, value="合并打印操作记录")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    ws.cell(row=1, column=1).font = openpyxl.styles.Font(bold=True, size=14)

    # ── 汇总信息 ──
    row = 3
    ws.cell(row=row, column=1, value="银行报盘文件数")
    ws.cell(row=row, column=2, value=len(renamed))
    row += 1
    ws.cell(row=row, column=1, value="合并文件数")
    ws.cell(row=row, column=2, value=len(matched))
    row += 1
    if success is not None:
        ws.cell(row=row, column=1, value="打印成功数")
        ws.cell(row=row, column=2, value=success)
        row += 1
        ws.cell(row=row, column=1, value="打印失败数")
        ws.cell(row=row, column=2, value=fail)
        row += 1
    ws.cell(row=row, column=1, value="未匹配到报盘的工资表数")
    ws.cell(row=row, column=2, value=len(unmatched))
    row += 1
    ws.cell(row=row, column=1, value="存在多个工资表的单位数")
    ws.cell(row=row, column=2, value=len(duplicates))

    # 建立原始报盘文件查找表：unit_name -> [old_name, ...]
    orig_files = {}
    for new_name, old_name, bank_name, unit_name in renamed:
        orig_files.setdefault(unit_name, []).append(old_name)

    # ── 明细表头 ──
    row += 2
    headers = ["序号", "单位名称", "合并文件", "原始报盘文件", "工资表文件", "匹配状态", "打印状态", "备注"]
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=ci, value=h)
        cell.font = openpyxl.styles.Font(bold=True)
        cell.fill = openpyxl.styles.PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # 建立打印结果查找表：unit_name -> "成功"/"失败"/"未打印"
    print_status = {}
    if success is not None:
        # 默认未打印
        for _, _, unit_name in matched:
            print_status[unit_name] = "未打印"
        # 覆盖成功/失败
        for merged_name, payroll_path, unit_name in matched:
            if payroll_path is None:
                print_status[unit_name] = "未打印（无工资表）"
        # 如果有 fail_list，标记失败
        for merged_name, unit_name in (fail_list or []):
            print_status[unit_name] = "失败"
        # 其余标记成功（排除了无工资表的和失败的）
        for merged_name, payroll_path, unit_name in matched:
            if payroll_path and print_status.get(unit_name) == "未打印":
                print_status[unit_name] = "成功"
    else:
        for _, _, unit_name in matched:
            print_status[unit_name] = "未打印"

    # ── 明细数据 ──
    for idx, (merged_name, payroll_path, unit_name) in enumerate(matched, 1):
        row += 1
        ws.cell(row=row, column=1, value=idx)
        ws.cell(row=row, column=2, value=unit_name)
        ws.cell(row=row, column=3, value=merged_name)
        # 原始报盘文件
        orig_list = orig_files.get(unit_name, [])
        ws.cell(row=row, column=4, value="\n".join(orig_list))
        ws.cell(row=row, column=5, value=os.path.basename(payroll_path) if payroll_path else "")
        ws.cell(row=row, column=6, value="已匹配" if payroll_path else "未找到")
        ws.cell(row=row, column=7, value=print_status.get(unit_name, ""))
        # 备注：重复文件提示
        note = ""
        if unit_name in duplicates:
            note = f"存在多个工资表文件，仅使用 {os.path.basename(payroll_path) if payroll_path else '第一个'}"
        ws.cell(row=row, column=8, value=note)

    # ── 金额警告 ──
    if merge_warnings:
        row += 2
        ws.cell(row=row, column=1, value="金额警告（分以下数值已舍去）：")
        ws.cell(row=row, column=1).font = openpyxl.styles.Font(bold=True, color="FF8C00")
        for w in merge_warnings:
            row += 1
            ws.cell(row=row, column=1, value=w)

    # ── 未匹配工资表 ──
    if unmatched:
        row += 2
        ws.cell(row=row, column=1, value="未匹配到银行报盘的工资表：")
        ws.cell(row=row, column=1).font = openpyxl.styles.Font(bold=True, color="FF0000")
        for fname in unmatched:
            row += 1
            ws.cell(row=row, column=1, value=fname)

    # 列宽
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 45
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].width = 45

    wb.save(report_path)
    return report_path


# ──────────────────────────────────────────────
# GUI 界面
# ──────────────────────────────────────────────


class BatchPrintGUI:
    """批量银行报盘合并与工资表打印工具 GUI"""

    def __init__(self, root):
        self.root = root
        self.root.title("批量打印程序")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        self.bank_dir = None
        self.payroll_dir = None
        self.output_dir = None

        self._build_ui()

    # ── UI 构建 ──────────────────────────────

    def _build_ui(self):
        # 标题
        title = tk.Label(
            self.root,
            text="批量银行报盘合并与工资表打印工具",
            font=("微软雅黑", 14, "bold"),
            pady=12,
        )
        title.pack(fill=tk.X)

        # ── 目录选择区域 ──
        dir_frame = tk.Frame(self.root, padx=12, pady=6)
        dir_frame.pack(fill=tk.X)

        tk.Button(dir_frame, text="选择银行报盘目录", command=self.select_bank_dir, width=18).grid(
            row=0, column=0, padx=(0, 8), pady=4
        )
        tk.Button(dir_frame, text="选择工资表目录", command=self.select_payroll_dir, width=18).grid(
            row=0, column=1, padx=8, pady=4
        )
        tk.Button(dir_frame, text="选择输出目录", command=self.select_output_dir, width=18).grid(
            row=0, column=2, padx=8, pady=4
        )

        # 路径显示
        self.bank_dir_label = tk.Label(dir_frame, text="银行报盘目录：未选择", anchor=tk.W, fg="#555")
        self.bank_dir_label.grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=2)

        self.payroll_dir_label = tk.Label(dir_frame, text="工资表目录：未选择", anchor=tk.W, fg="#555")
        self.payroll_dir_label.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=2)

        self.output_dir_label = tk.Label(dir_frame, text="输出目录：未选择", anchor=tk.W, fg="#555")
        self.output_dir_label.grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=2)

        dir_frame.columnconfigure(0, weight=1)
        dir_frame.columnconfigure(1, weight=1)
        dir_frame.columnconfigure(2, weight=1)

        # ── 按钮区域（目录选择和日志之间） ──
        btn_frame = tk.Frame(self.root, padx=12, pady=10)
        btn_frame.pack(fill=tk.X, pady=(4, 0))

        self.run_all_btn = tk.Button(
            btn_frame,
            text="执行全部并打印",
            command=self.run_all_and_print,
            bg="#4a90d9",
            fg="white",
            font=("微软雅黑", 11, "bold"),
            padx=20,
            pady=4,
        )
        self.run_all_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.merge_only_btn = tk.Button(
            btn_frame,
            text="仅合并不打印",
            command=self.run_merge_only,
            bg="#6abf69",
            fg="white",
            font=("微软雅黑", 11, "bold"),
            padx=20,
            pady=4,
        )
        self.merge_only_btn.pack(side=tk.LEFT)

        # 分隔线
        sep = tk.Frame(self.root, height=2, bd=1, relief=tk.SUNKEN)
        sep.pack(fill=tk.X, padx=10, pady=6)

        # ── 日志区域 ──
        log_frame = tk.Frame(self.root, padx=12)
        log_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(log_frame, text="日志信息", font=("微软雅黑", 10, "bold"), anchor=tk.W).pack(
            fill=tk.X, pady=(0, 4)
        )

        self.log_area = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg="#fafafa",
            state=tk.DISABLED,
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

    # ── 日志 ──────────────────────────────

    def log(self, message):
        """向日志区域追加一行"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.root.update_idletasks()

    # ── 目录选择 ──────────────────────────────

    def _set_default_output_dir(self):
        """根据银行报盘目录自动设置输出目录（同级目录）"""
        if not self.bank_dir:
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        parent = os.path.dirname(self.bank_dir)
        default_dir = os.path.join(parent, f"合并后的银行报盘_{ts}")
        self.output_dir = default_dir
        self.output_dir_label.config(text=f"输出目录：{default_dir}")

    def select_bank_dir(self):
        d = filedialog.askdirectory(title="选择银行报盘目录")
        if d:
            self.bank_dir = d
            self.bank_dir_label.config(text=f"银行报盘目录：{d}")
            self._set_default_output_dir()

    def select_payroll_dir(self):
        d = filedialog.askdirectory(title="选择工资表目录")
        if d:
            self.payroll_dir = d
            self.payroll_dir_label.config(text=f"工资表目录：{d}")

    def select_output_dir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.output_dir = d
            self.output_dir_label.config(text=f"输出目录：{d}")

    # ── 进度控制 ──────────────────────────────

    def _set_busy(self, busy):
        """切换按钮启用状态"""
        state = tk.DISABLED if busy else tk.NORMAL
        self.run_all_btn.config(state=state)
        self.merge_only_btn.config(state=state)
        self.root.update_idletasks()

    # ── 执行全部并打印 ──────────────────────────────

    def run_all_and_print(self):
        if not self._check_dirs():
            return

        self._set_busy(True)
        self.log("=" * 50)
        self.log("开始执行全部流程...")
        self.log("")

        # 步骤 1：改名银行报盘文件
        self.log("【步骤1】改名银行报盘文件...")
        try:
            renamed = rename_bank_files(self.bank_dir, self.output_dir)
            self.log(f"  ✓ 改名完成：{len(renamed)} 个文件")
        except Exception as e:
            self.log(f"  ✗ 改名失败：{e}")
            self._set_busy(False)
            return

        # 步骤 2：合并银行报盘文件
        self.log("【步骤2】合并银行报盘文件...")
        try:
            merged, merge_warnings = merge_bank_files(renamed, self.bank_dir, self.output_dir)
            self.log(f"  ✓ 合并完成：{len(merged)} 个合并文件")
            if merge_warnings:
                self.log(f"  ⚠ 金额警告（分以下数值已舍去）：")
                for w in merge_warnings:
                    self.log(f"    - {w}")
        except Exception as e:
            self.log(f"  ✗ 合并失败：{e}")
            self._set_busy(False)
            return

        # 步骤 3：匹配工资表
        self.log("【步骤3】匹配工资表文件...")
        try:
            matched, unmatched, duplicates = match_payroll_files(merged, self.payroll_dir)
            matched_count = sum(1 for _, fp, _ in matched if fp is not None)
            self.log(f"  ✓ 匹配完成：{matched_count}/{len(matched)} 匹配成功")
            for merged_name, payroll_path, unit_name in matched:
                status = f"→ {payroll_path}" if payroll_path else "✗ 未找到匹配"
                self.log(f"    {unit_name}: {status}")
            if duplicates:
                self.log(f"  ⚠ 以下单位存在多个工资表文件，仅使用第一个：")
                for unit_name, files in duplicates.items():
                    self.log(f"    {unit_name}:")
                    for f in files:
                        self.log(f"      - {f}")
            if unmatched:
                self.log(f"  ⚠ 以下 {len(unmatched)} 个工资表未匹配到银行报盘，不参与打印：")
                for fname in unmatched:
                    self.log(f"    - {fname}")
        except Exception as e:
            self.log(f"  ✗ 匹配失败：{e}")
            self._set_busy(False)
            return

        # 步骤 4：打印（需要用户确认）
        # 文件名已带序号前缀，字典序即正确顺序
        matched.sort(key=lambda x: x[0])
        self.log("")
        self.log("【步骤4】准备打印...")

        if not check_wps_available():
            self.log("  ✗ WPS Office 不可用，无法打印")
            self.log("  提示：打印功能需要 Windows 环境 + WPS Office")
            self._set_busy(False)
            return

        # 展示合并文件列表供用户确认
        self.log("")
        self.log("以下文件将打印：")
        for merged_name, payroll_path, unit_name in matched:
            if payroll_path:
                self.log(f"  • {unit_name} → {os.path.basename(payroll_path)}")

        self.log("")
        self.log("请在弹出对话框确认后开始打印...")
        self.root.update()

        # 弹出确认对话框
        ok = messagebox.askyesno(
            "确认打印",
            f"将打印 {matched_count} 个单位的工资表，是否继续？",
        )
        if not ok:
            self.log("  用户取消打印")
            self._set_busy(False)
            return

        # 执行打印
        self.log("")
        self.log("开始打印...")

        def progress_cb(current, total, message):
            self.log(f"  [{current}/{total}] {message}")
            self.root.update()

        success, fail, fail_list = batch_print(matched, progress_cb)

        self.log("")
        self.log(f"打印完成：成功 {success}，失败 {fail}")
        if fail_list:
            self.log("失败列表：")
            for merged_name, unit_name in fail_list:
                self.log(f"  ✗ {unit_name} ({merged_name})")

        # 生成操作记录
        try:
            report_path = generate_report_xlsx(
                self.output_dir, renamed, matched,
                unmatched, duplicates, merge_warnings, success, fail, fail_list
            )
            self.log(f"  📄 操作记录已保存：{report_path}")
        except Exception as e:
            self.log(f"  ✗ 操作记录生成失败：{e}")

        self.log("=" * 50)
        self._set_busy(False)

    # ── 仅合并不打印 ──────────────────────────────

    def run_merge_only(self):
        if not self._check_dirs():
            return

        self._set_busy(True)
        self.log("=" * 50)
        self.log("开始执行合并流程（不打印）...")
        self.log("")

        # 步骤 1：改名
        self.log("【步骤1】改名银行报盘文件...")
        try:
            renamed = rename_bank_files(self.bank_dir, self.output_dir)
            self.log(f"  ✓ 改名完成：{len(renamed)} 个文件")
        except Exception as e:
            self.log(f"  ✗ 改名失败：{e}")
            self._set_busy(False)
            return

        # 步骤 2：合并
        self.log("【步骤2】合并银行报盘文件...")
        try:
            merged, merge_warnings = merge_bank_files(renamed, self.bank_dir, self.output_dir)
            self.log(f"  ✓ 合并完成：{len(merged)} 个合并文件")
            if merge_warnings:
                self.log(f"  ⚠ 金额警告（分以下数值已舍去）：")
                for w in merge_warnings:
                    self.log(f"    - {w}")
        except Exception as e:
            self.log(f"  ✗ 合并失败：{e}")
            self._set_busy(False)
            return

        # 步骤 3：匹配
        self.log("【步骤3】匹配工资表文件...")
        try:
            matched, unmatched, duplicates = match_payroll_files(merged, self.payroll_dir)
            # 文件名已带序号前缀，字典序即正确顺序
            matched.sort(key=lambda x: x[0])
            matched_count = sum(1 for _, fp, _ in matched if fp is not None)
            self.log(f"  ✓ 匹配完成：{matched_count}/{len(matched)} 匹配成功")
            for merged_name, payroll_path, unit_name in matched:
                status = f"→ {payroll_path}" if payroll_path else "✗ 未找到匹配"
                self.log(f"    {unit_name}: {status}")
            if duplicates:
                self.log(f"  ⚠ 以下单位存在多个工资表文件，仅使用第一个：")
                for unit_name, files in duplicates.items():
                    self.log(f"    {unit_name}:")
                    for f in files:
                        self.log(f"      - {f}")
            if unmatched:
                self.log(f"  ⚠ 以下 {len(unmatched)} 个工资表未匹配到银行报盘，不参与打印：")
                for fname in unmatched:
                    self.log(f"    - {fname}")
        except Exception as e:
            self.log(f"  ✗ 匹配失败：{e}")
            self._set_busy(False)
            return

        self.log("")
        self.log("合并流程完成，未执行打印。")

        # 生成操作记录
        try:
            report_path = generate_report_xlsx(
                self.output_dir, renamed, matched, unmatched, duplicates, merge_warnings
            )
            self.log(f"  📄 操作记录已保存：{report_path}")
        except Exception as e:
            self.log(f"  ✗ 操作记录生成失败：{e}")

        self.log("=" * 50)
        self._set_busy(False)

    # ── 目录检查 ──────────────────────────────

    def _check_dirs(self):
        """检查三个目录是否都已选择"""
        if not self.bank_dir:
            messagebox.showwarning("提示", "请先选择银行报盘目录")
            return False
        if not self.payroll_dir:
            messagebox.showwarning("提示", "请先选择工资表目录")
            return False
        if not self.output_dir:
            messagebox.showwarning("提示", "请先选择输出目录")
            return False
        return True


def main():
    """启动 GUI"""
    root = tk.Tk()
    app = BatchPrintGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

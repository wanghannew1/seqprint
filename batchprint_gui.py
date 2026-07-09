"""
批量银行报盘合并与工资表打印工具
"""

import os
import shutil
import tkinter as tk
from datetime import datetime
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP
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
        rows.append([seq, account, name, amount, "", "", "", "", ""])
    return rows


def _read_ccb_rows(filepath, warnings):
    """读取建设银行 9 列格式，跨行标识和行名不填，建行系统自动识别"""
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
        # 跨行标识和行名不填，建行系统自动识别
        row[4] = ""
        row[5] = ""
        rows.append(row)
    return rows


def _read_jlb_rows(filepath, warnings):
    """读取吉林银行 9 列格式，跨行标识和行名不填，建行系统自动识别"""
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
        # 跨行标识和行名不填，建行系统自动识别
        row[4] = ""
        row[5] = ""
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

        # 写入（建行系统兼容：空值列不写入单元格，金额列设小数格式）
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = unit_name
        ws.append(HEADERS)
        for row_idx, row in enumerate(all_rows, start=2):
            for c, val in enumerate(row, start=1):
                if val is None or val == "":
                    continue  # 跳过空值，建行系统不接受空 inlineStr 单元格
                cell = ws.cell(row=row_idx, column=c, value=val)
                if c == 4:  # 金额列设两位小数格式
                    cell.number_format = "0.00"
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
# 工资表合并（按个税分组）
# ──────────────────────────────────────────────


def _read_payroll_data(filepath):
    """读取工资表文件，返回 (header_rows, data_rows, footer_rows, tax_col_idx)
    header_rows: 表头行列表（每行是单元格值列表）
    data_rows: 数据行列表（每行是单元格值列表）
    footer_rows: 表尾行列表（备注行、签字行等）
    tax_col_idx: 个人所得税列索引（0-based）
    """
    is_xls = filepath.lower().endswith(".xls")
    if is_xls:
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_index(0)
        nrows, ncols = ws.nrows, ws.ncols
    else:
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        nrows, ncols = ws.max_row, ws.max_column

    # 找个人所得税列
    tax_col = None
    for c in range(ncols):
        if is_xls:
            val = str(ws.cell_value(2, c)).strip()
        else:
            val = str(ws.cell(row=3, column=c + 1).value or "").strip()
        if "个人所得" in val or "个税" in val:
            tax_col = c
            break

    # 读取所有行
    all_rows = []
    for r in range(nrows):
        row = []
        for c in range(ncols):
            if is_xls:
                row.append(ws.cell_value(r, c))
            else:
                row.append(ws.cell(row=r + 1, column=c + 1).value)
        all_rows.append(row)

    # 分离表头、数据、表尾
    # 表头：行0-4（标题行、单位行、列名行、扣款子行1、扣款子行2）
    # 数据行：有序号（第1列是数字）且姓名列有值且不是"合计"的行
    # 表尾：其余行（备注行、签字行、空行）
    header_rows = all_rows[:5]
    data_rows = []
    footer_rows = []
    for row in all_rows[5:]:
        seq = str(row[0]).strip() if len(row) > 0 else ""
        name = str(row[1]).strip() if len(row) > 1 else ""
        # 数据行特征：序号是数字，姓名有值且不是合计/转账等汇总行
        is_data = False
        try:
            float(seq)  # 序号是数字
            if name and name not in ("合计", "转账合计"):
                is_data = True
        except (ValueError, TypeError):
            pass
        if is_data:
            data_rows.append(row)
        else:
            footer_rows.append(row)

    return header_rows, data_rows, footer_rows, tax_col


def _get_common_columns(all_headers):
    """从所有表头中找出共同列和独有列
    返回: (common_cols, unique_cols_map)
    common_cols: [(idx, name), ...] 所有文件都有的列
    unique_cols_map: {filename: [(idx, name), ...]} 每个文件独有的列
    """
    # 取第一个文件的列名作为基准
    base = all_headers[0][2]  # 第3行是列名行
    base_names = [str(v)[:20] for v in base]

    common = list(range(len(base_names)))
    unique_map = {}

    for i, headers in enumerate(all_headers[1:], 1):
        names = [str(v)[:20] for v in headers[2]]
        # 找当前文件比基准多的列
        extra = []
        for c, name in enumerate(names):
            if name and name not in base_names:
                extra.append((c, name))
        if extra:
            unique_map[i] = extra

    return [(i, base_names[i]) for i in range(len(base_names))], unique_map


def _detect_column_structure(headers):
    """检测工资表的列结构类型，返回列名列表和特殊列信息"""
    row3 = [str(v or "").strip() for v in headers[2]]
    row4 = [str(v or "").strip() for v in headers[3]] if len(headers) > 3 else []
    row5 = [str(v or "").strip() for v in headers[4]] if len(headers) > 4 else []

    # 构建列名（合并行3-5的信息）
    col_names = []
    for c in range(len(row3)):
        name = row3[c]
        # 扣款明细主列名
        if "扣" in name and "款" in name:
            name = "扣款明细"
        col_names.append(name)

    # 检测特殊列
    special_cols = {}
    for c, name in enumerate(col_names):
        if "补发工资" in name:
            special_cols["补发工资"] = c
        elif "大病险" in name:
            special_cols["大病险"] = c
        elif "雇主责任险" in name:
            special_cols["雇主责任险"] = c
        elif "公务员医疗补助" in name:
            special_cols["公务员医疗补助"] = c
        elif "单位缴纳五险一金" in name:
            special_cols["单位缴纳五险一金"] = c

    # 检查行4/5中是否有公务员医疗补助
    for r in [row4, row5]:
        for c, v in enumerate(r):
            if "公务员医疗补助" in v:
                special_cols["公务员医疗补助"] = c

    return col_names, special_cols


def _build_universal_column_map(col_names, special_cols):
    """将任意列结构映射到统一输出列（删除部门/岗位/职工号，保留特殊列）
    返回: (output_cols, delete_indices)
    output_cols: 输出列索引列表（0-based），每个元素为 (输出列索引, 原始列索引)
    delete_indices: 被删除的原始列索引
    """
    # 标准列顺序（所有文件都有）
    # 序号(0), 姓名(1), 身份证(2), 部门(3), 岗位(4), 职工号(5),
    # 基本工资(6), 交通补贴(7) / 补发工资(7), 应发工资(8/9),
    # 单位缴纳五险一金(9/10), 单位代理费(10/11), 大病险(11/12),
    # 雇主责任险(11/12), 转账合计(11/12/13),
    # 社保基数(12/13/14), 医保基数(13/14/15), 工伤基数(14/15/16), 公积金基数(15/16/17),
    # 扣款明细(16-26/17-27), 个人所得税(27/28), 实发工资(28/29), 实发合计(29/30)

    # 删除列：部门(3), 岗位(4), 职工号(5)
    delete_indices = {3, 4, 5}

    # 构建输出映射
    output_cols = []
    for c in range(len(col_names)):
        if c in delete_indices:
            continue
        output_cols.append(c)

    return output_cols, delete_indices


def _read_payroll_workbook(filepath):
    """读取工资表文件，返回 (wb_or_ws, is_xls, nrows, ncols, merged_cells, images)"""
    is_xls = filepath.lower().endswith(".xls")
    if is_xls:
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_index(0)
        return (wb, ws, is_xls, ws.nrows, ws.ncols, [], [])
    else:
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        merged = list(ws.merged_cells.ranges) if ws.merged_cells else []
        # 从 xlsx zip 中直接读取图片文件（openpyxl 的 img.ref 数据被压缩过，不能直接用）
        images = []
        if ws._images:
            import zipfile
            try:
                with zipfile.ZipFile(filepath) as z:
                    for img in ws._images:
                        try:
                            png_path = img.path.lstrip('/')
                            png_data = z.read(png_path)
                            images.append((png_data, img.anchor._from.row, img.anchor._from.col, img.width, img.height))
                        except Exception:
                            pass
            except Exception:
                pass
        return (wb, ws, is_xls, ws.max_row, ws.max_column, merged, images)


def _build_column_name_map(headers):
    """从表头构建列名→索引映射。
    Merged cells in row3/row4 produce empty subsequent entries; we track
    the last non-empty row3 and row4 values to create unique keys.
    返回 (name_to_idx, col_names):
    name_to_idx: {规范列名: 0-based索引}
    col_names: [规范列名, ...] 按原始列顺序
    """
    row3 = [str(v or "").strip() for v in headers[2]]
    row4 = [str(v or "").strip() for v in headers[3]] if len(headers) > 3 else []
    row5 = [str(v or "").strip() for v in headers[4]] if len(headers) > 4 else []

    name_to_idx = {}
    col_names = []
    last_row3 = ""
    last_row4 = ""

    for c in range(len(row3)):
        raw = row3[c]
        if raw and "扣" in raw and "款" in raw:
            name = "扣款明细"
        elif raw:
            name = raw
        else:
            name = last_row3
        if raw:
            last_row3 = name

        # Propagate row4 only within merged ranges (row3 empty)
        r4_val = ""
        if c < len(row4) and row4[c]:
            r4_val = row4[c]
            last_row4 = row4[c]
        elif c < len(row4) and not raw and last_row4:
            r4_val = last_row4
        elif c < len(row4) and raw:
            last_row4 = ""  # Reset when entering a new section

        sub_parts = [r4_val] if r4_val else []
        if c < len(row5) and row5[c]:
            sub_parts.append(row5[c])
        sub_name = "/".join(sub_parts) if sub_parts else ""

        if sub_name:
            key = f"{name}/{sub_name}" if name != sub_name else name
        else:
            key = name
        name_to_idx[key] = c
        col_names.append(key)

    return name_to_idx, col_names


def _get_canonical_columns(all_file_columns):
    """从所有文件的列名构建规范列顺序。
    使用列数最多的文件作为基准（包含所有可选列），其他文件按名称补齐。
    返回: [规范列名, ...] 不包含 部门/岗位/职工号
    """
    # 找到列数最多的文件作为规范基准
    longest = max(all_file_columns, key=len)
    canonical = [c for c in longest if c not in ("部门", "岗位", "职工号")]
    return canonical


def _normalize_row_by_names(row, headers, canonical_cols):
    """根据规范列名将原始行映射为规范行。
    row: 原始数据行（值列表）
    headers: 该文件的表头
    canonical_cols: 规范列名列表
    返回: [值, ...] 按规范列顺序
    """
    name_to_idx, _ = _build_column_name_map(headers)
    result = []
    for cname in canonical_cols:
        if cname in name_to_idx:
            idx = name_to_idx[cname]
            if idx < len(row):
                result.append(row[idx])
            else:
                result.append("")
        else:
            result.append("")
    return result


def _adjust_merged_range(mrange, delete_1based_set, insert_before_1based=3):
    """调整合并单元格范围，处理列删除和插入。
    mrange: openpyxl merged cell range 或 str
    delete_1based_set: 被删除列的1-based列号集合
    insert_before_1based: 在1-based列号前插入新列（结算单元，位于姓名后=原C前）
    返回: 调整后的范围字符串，或 None
    """
    import re
    rng_str = str(mrange)
    parts = rng_str.split(":")
    if len(parts) != 2:
        return None

    from_cell, to_cell = parts[0], parts[1]
    m1 = re.match(r"([A-Z]+)(\d+)", from_cell)
    m2 = re.match(r"([A-Z]+)(\d+)", to_cell)
    if not m1 or not m2:
        return None

    from_col_str, from_row = m1.group(1), int(m1.group(2))
    to_col_str, to_row = m2.group(1), int(m2.group(2))

    from_col = sum((ord(c) - 64) * (26 ** i) for i, c in enumerate(reversed(from_col_str)))
    to_col = sum((ord(c) - 64) * (26 ** i) for i, c in enumerate(reversed(to_col_str)))

    def _mapped_col(col_1based):
        """Original 1-based → output 1-based, accounting for deletion then insertion."""
        # Step 1: deletion (3 cols removed at 1-based 4,5,6)
        deleted_before = sum(1 for d in sorted(delete_1based_set) if d < col_1based)
        after_delete = col_1based - deleted_before
        if after_delete <= 0:
            return None
        # Step 2: insertion (结算单元 inserted at 1-based column = insert_before_1based)
        if after_delete >= insert_before_1based:
            return after_delete + 1
        else:
            return after_delete

    new_from = _mapped_col(from_col)
    new_to = _mapped_col(to_col)

    if new_from is None or new_to is None or new_from > new_to:
        return None

    def col_to_letter(n):
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    return f"{col_to_letter(new_from)}{from_row}:{col_to_letter(new_to)}{to_row}"


def merge_payrolls_by_tax(payroll_dir, output_dir, bank_dir=None):
    """
    合并所有工资表，按个人所得税>0和=0分成两张表
    同时按分组生成对应的银行报盘文件
    返回: (tax_path, no_tax_path, tax_bank_path, no_tax_bank_path)
    """
    os.makedirs(output_dir, exist_ok=True)

    # 扫描工资表
    payroll_files = [
        f for f in os.listdir(payroll_dir)
        if "汇总表" not in f and "验证" not in f
        and f.endswith((".xlsx", ".xls"))
    ]

    # 去重排序：同一单位优先保留 signed_ 版本，其次 xlsx，最后 xls
    unit_file_map = {}
    for fname in payroll_files:
        unit_name = fname
        for sep in ["2026年06月工资表", "202606工资表", "202606工资"]:
            if sep in fname:
                unit_name = fname.split(sep, 1)[0].strip()
                break
        unit_name = unit_name.replace("signed_", "").strip()
        is_signed = fname.startswith("signed_")
        is_xlsx = fname.endswith(".xlsx")
        priority = 0 if is_signed and is_xlsx else 1 if is_signed else 2 if is_xlsx else 3
        if unit_name not in unit_file_map or priority < unit_file_map[unit_name][0]:
            unit_file_map[unit_name] = (priority, fname)

    sorted_files = sorted(unit_file_map.values(), key=lambda x: x[0])

    # 读取所有工资表
    all_data = []
    all_file_col_names = []
    file_name_to_idx_map = {}
    sample_header = None
    sample_merged_cells = None
    sample_images = None

    for priority, fname in sorted_files:
        path = os.path.join(payroll_dir, fname)
        headers, data_rows, footers, tax_col = _read_payroll_data(path)

        name_to_idx, col_names = _build_column_name_map(headers)
        all_file_col_names.append(col_names)
        file_name_to_idx_map[fname] = name_to_idx

        if sample_header is None:
            sample_header = headers
            if fname.startswith("signed_") and fname.endswith(".xlsx"):
                try:
                    _, ws, _, _, _, merged, images = _read_payroll_workbook(path)
                    sample_merged_cells = merged
                    sample_images = images
                except Exception:
                    pass

        unit_name = fname
        for sep in ["2026年06月工资表", "202606工资表", "202606工资"]:
            if sep in fname:
                unit_name = fname.split(sep, 1)[0].strip()
                break
        unit_name = unit_name.replace("signed_", "").strip()

        for row in data_rows:
            tax_val = row[tax_col] if tax_col is not None else 0
            try:
                tax_amt = float(tax_val) if tax_val else 0
            except (ValueError, TypeError):
                tax_amt = 0
            all_data.append((unit_name, tax_amt, row, headers))

    if not all_data:
        return None, None, None, None

    canonical_cols = _get_canonical_columns(all_file_col_names)
    canonical_cols = [c for c in canonical_cols if c not in ("部门", "岗位", "职工号")]

    normalized_data = []
    for unit_name, tax_amt, row, headers in all_data:
        norm_row = _normalize_row_by_names(row, headers, canonical_cols)
        normalized_data.append((unit_name, tax_amt, norm_row))

    name_col_in_canonical = None
    for c, name in enumerate(canonical_cols):
        if name == "姓名":
            name_col_in_canonical = c
            break
    if name_col_in_canonical is not None:
        insert_pos = name_col_in_canonical + 1
        canonical_cols.insert(insert_pos, "结算单元")
        for i, (uname, tax_amt, nrow) in enumerate(normalized_data):
            nrow.insert(insert_pos, uname)

    tax_group = [d for d in normalized_data if d[1] > 0]
    no_tax_group = [d for d in normalized_data if d[1] == 0]

    max_output_cols = len(canonical_cols)

    header_rows = []
    title_row = [""] * max_output_cols
    title_row[0] = "吉林大学2026年06月人才派遣人员工资发放表"
    header_rows.append(title_row)

    unit_row = [""] * max_output_cols
    unit_row[0] = "单位"
    unit_row[1] = "名称：吉林大学"
    header_rows.append(unit_row)

    main_row = []
    sub1_row = []
    sub2_row = []
    prev_main = ""
    prev_sub1 = ""
    for cname in canonical_cols:
        if cname == "结算单元":
            main_row.append("结算单元")
            sub1_row.append("")
            sub2_row.append("")
            prev_main = "结算单元"
            prev_sub1 = ""
            continue
        if "/" in cname:
            parts = cname.split("/", 1)
            main_val = parts[0]
            if main_val == prev_main:
                main_row.append("")
            else:
                main_row.append(main_val)
                prev_main = main_val
            sub = parts[1]
            if "/" in sub:
                sp = sub.split("/", 1)
                sub1_val = sp[0]
                if sub1_val == prev_sub1:
                    sub1_row.append("")
                else:
                    sub1_row.append(sub1_val)
                    prev_sub1 = sub1_val
                sub2_row.append(sp[1])
            else:
                if sub == prev_sub1:
                    sub1_row.append("")
                else:
                    sub1_row.append(sub)
                    prev_sub1 = sub
                sub2_row.append("")
        else:
            if cname == prev_main:
                main_row.append("")
            else:
                main_row.append(cname)
                prev_main = cname
            sub1_row.append("")
            sub2_row.append("")

    header_rows.append(main_row)
    if any(v for v in sub1_row):
        header_rows.append(sub1_row)
    if any(v for v in sub2_row):
        header_rows.append(sub2_row)

    detail_start = None
    detail_end = None
    for c, cname in enumerate(canonical_cols, 1):
        if cname.startswith("扣款明细"):
            if detail_start is None:
                detail_start = c
            detail_end = c
    title_end_col = openpyxl.utils.get_column_letter(max_output_cols)
    detail_start_col = openpyxl.utils.get_column_letter(detail_start) if detail_start else "A"
    detail_end_col = openpyxl.utils.get_column_letter(detail_end) if detail_end else "A"

    canonical_merged = []
    canonical_merged.append(f"A1:{title_end_col}1")
    if detail_start and detail_end and detail_start < detail_end:
        canonical_merged.append(f"{detail_start_col}3:{detail_end_col}3")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    def write_payroll(group, suffix):
        if not group:
            return None
        fname = f"吉林大学2026年06月人才派遣人员工资发放表_{suffix}.xlsx"
        fpath = os.path.join(output_dir, fname)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"工资发放表_{suffix}"

        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins.left = 1.5
        ws.page_margins.right = 0.5
        ws.page_margins.top = 0.75
        ws.page_margins.bottom = 0.75

        for r_idx, row in enumerate(header_rows):
            for c_idx, val in enumerate(row):
                ws.cell(row=r_idx + 1, column=c_idx + 1, value=val)

        for i, (uname, tax_amt, row) in enumerate(group, 1):
            row[0] = i
            for c_idx, val in enumerate(row):
                ws.cell(row=len(header_rows) + i, column=c_idx + 1, value=val)
                if c_idx >= 6:
                    try:
                        float(val)
                        ws.cell(row=len(header_rows) + i, column=c_idx + 1).number_format = "0.00"
                    except (ValueError, TypeError):
                        pass

        total_row_idx = len(header_rows) + len(group) + 1
        ws.cell(row=total_row_idx, column=1, value="合计")
        for c in range(2, max_output_cols + 1):
            total = 0.0
            all_num = True
            for _, _, row in group:
                v = row[c - 1] if c - 1 < len(row) else 0
                try:
                    total += float(v or 0)
                except (ValueError, TypeError):
                    all_num = False
                    break
            if all_num:
                cell = ws.cell(row=total_row_idx, column=c, value=total)
                cell.number_format = "0.00"

        sign_row_idx = total_row_idx + 2
        sign_labels = {
            1: "总经理签字：",
            7: "部长签字：",
            13: "财务审核：",
            19: "制表人：张朦",
        }
        for col, label in sign_labels.items():
            ws.cell(row=sign_row_idx, column=col, value=label)

        for merge_range in canonical_merged:
            try:
                ws.merge_cells(merge_range)
            except Exception:
                pass

        if sample_images:
            from openpyxl.drawing.image import Image as XLImage
            from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
            import io
            # Map each image to a signature column (0-based)
            # sign_labels are at 1-based cols: 1(总经理), 7(部长), 13(财务), 19(制表人)
            sign_cols_0based = [0, 6, 12, 18]
            for i, (img_data, orig_row, orig_col, width, height) in enumerate(sample_images[:4]):
                try:
                    new_img = XLImage(io.BytesIO(img_data))
                    new_img.width = width
                    new_img.height = height
                    col = sign_cols_0based[i] if i < len(sign_cols_0based) else 0
                    new_anchor = OneCellAnchor()
                    new_anchor._from = AnchorMarker(col=col, row=sign_row_idx - 1)
                    new_img.anchor = new_anchor
                    ws.add_image(new_img)
                except Exception:
                    pass

        wb.save(fpath)
        return fpath

    tax_path = write_payroll(tax_group, "有个税")
    no_tax_path = write_payroll(no_tax_group, "无个税")

    # 生成对应的银行报盘文件
    tax_bank_path = None
    no_tax_bank_path = None
    if bank_dir:
        bank_files = [f for f in os.listdir(bank_dir) if f.lower().endswith(".xls")]
        bank_map = {}
        for fname in bank_files:
            try:
                _, _, unit_name = split_filename(fname)
                bank_map[unit_name] = fname
            except ValueError:
                continue

        def write_bank(group_data, suffix):
            if not group_data:
                return None
            all_bank_rows = []
            seen_units = set()
            for unit_name, _, _ in group_data:
                if unit_name in bank_map and unit_name not in seen_units:
                    seen_units.add(unit_name)
                    fname = bank_map[unit_name]
                    fpath = os.path.join(bank_dir, fname)
                    bt = detect_bank_type(fname)
                    if bt == "icbc":
                        rows = _read_icbc_rows(fpath, [])
                    elif bt == "ccb":
                        rows = _read_ccb_rows(fpath, [])
                    elif bt == "jlb":
                        rows = _read_jlb_rows(fpath, [])
                    all_bank_rows.extend(rows)

            if not all_bank_rows:
                return None

            for i, row in enumerate(all_bank_rows, 1):
                row[0] = i

            fname = f"银行报盘_{suffix}_{ts}.xlsx"
            fpath = os.path.join(output_dir, fname)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"银行报盘_{suffix}"
            ws.append(HEADERS)
            for row_idx, row in enumerate(all_bank_rows, start=2):
                for c, val in enumerate(row, start=1):
                    if val is None or val == "":
                        continue
                    cell = ws.cell(row=row_idx, column=c, value=val)
                    if c == 4:
                        cell.number_format = "0.00"
            wb.save(fpath)
            return fpath

        tax_bank_path = write_bank(tax_group, "有个税")
        no_tax_bank_path = write_bank(no_tax_group, "无个税")

    return tax_path, no_tax_path, tax_bank_path, no_tax_bank_path


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

        # ── 新功能：合并工资表及报盘 ──
        sep2 = tk.Frame(btn_frame, width=2, bd=1, relief=tk.SUNKEN, height=30)
        sep2.pack(side=tk.LEFT, padx=10)

        self.merge_payroll_btn = tk.Button(
            btn_frame,
            text="合并工资表及报盘并打印",
            command=self.run_merge_payroll_and_print,
            bg="#e67e22",
            fg="white",
            font=("微软雅黑", 11, "bold"),
            padx=16,
            pady=4,
        )
        self.merge_payroll_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.merge_payroll_no_print_btn = tk.Button(
            btn_frame,
            text="合并工资表及报盘不打印",
            command=self.run_merge_payroll_only,
            bg="#9b59b6",
            fg="white",
            font=("微软雅黑", 11, "bold"),
            padx=16,
            pady=4,
        )
        self.merge_payroll_no_print_btn.pack(side=tk.LEFT)

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
        self.merge_payroll_btn.config(state=state)
        self.merge_payroll_no_print_btn.config(state=state)
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

        # 仅打印匹配到工资表的单位
        matched_to_print = [(m, p, u) for m, p, u in matched if p is not None]
        self.log(f"  实际需打印：{len(matched_to_print)} 个（未匹配的 {len(matched) - len(matched_to_print)} 个已跳过）")
        success, fail, fail_list = batch_print(matched_to_print, progress_cb)

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

    # ── 合并工资表及报盘并打印 ──────────────────────────────

    def run_merge_payroll_and_print(self):
        if not self._check_dirs():
            return

        self._set_busy(True)
        self.log("=" * 50)
        self.log("开始合并工资表及报盘（按个税分组）...")
        self.log("")

        try:
            tax_path, no_tax_path, tax_bank_path, no_tax_bank_path = merge_payrolls_by_tax(
                self.payroll_dir, self.output_dir, self.bank_dir
            )
        except Exception as e:
            self.log(f"  ✗ 合并失败：{e}")
            import traceback
            self.log(traceback.format_exc())
            self._set_busy(False)
            return

        self.log(f"  📁 输出目录：{self.output_dir}")
        if tax_path:
            self.log(f"  ✓ 有个税工资表：{os.path.basename(tax_path)}")
        if no_tax_path:
            self.log(f"  ✓ 无个税工资表：{os.path.basename(no_tax_path)}")
        if tax_bank_path:
            self.log(f"  ✓ 有个税银行报盘：{os.path.basename(tax_bank_path)}")
        if no_tax_bank_path:
            self.log(f"  ✓ 无个税银行报盘：{os.path.basename(no_tax_bank_path)}")

        # 打印
        self.log("")
        self.log("准备打印...")
        if not check_wps_available():
            self.log("  ✗ WPS Office 不可用，无法打印")
            self._set_busy(False)
            return

        to_print = []
        if tax_path:
            to_print.append(("", tax_path, "有个税工资表"))
        if no_tax_path:
            to_print.append(("", no_tax_path, "无个税工资表"))

        self.log("以下文件将打印：")
        for _, fp, label in to_print:
            self.log(f"  • {label} → {os.path.basename(fp)}")

        ok = messagebox.askyesno("确认打印", f"将打印 {len(to_print)} 张工资表，是否继续？")
        if not ok:
            self.log("  用户取消打印")
            self._set_busy(False)
            return

        def progress_cb(current, total, message):
            self.log(f"  [{current}/{total}] {message}")
            self.root.update()

        success, fail, fail_list = batch_print(to_print, progress_cb)
        self.log(f"打印完成：成功 {success}，失败 {fail}")

        self.log("=" * 50)
        self._set_busy(False)

    # ── 合并工资表及报盘不打印 ──────────────────────────────

    def run_merge_payroll_only(self):
        if not self._check_dirs():
            return

        self._set_busy(True)
        self.log("=" * 50)
        self.log("开始合并工资表及报盘（按个税分组，不打印）...")
        self.log("")

        try:
            tax_path, no_tax_path, tax_bank_path, no_tax_bank_path = merge_payrolls_by_tax(
                self.payroll_dir, self.output_dir, self.bank_dir
            )
        except Exception as e:
            self.log(f"  ✗ 合并失败：{e}")
            import traceback
            self.log(traceback.format_exc())
            self._set_busy(False)
            return

        self.log(f"  📁 输出目录：{self.output_dir}")
        if tax_path:
            self.log(f"  ✓ 有个税工资表：{os.path.basename(tax_path)}")
        if no_tax_path:
            self.log(f"  ✓ 无个税工资表：{os.path.basename(no_tax_path)}")
        if tax_bank_path:
            self.log(f"  ✓ 有个税银行报盘：{os.path.basename(tax_bank_path)}")
        if no_tax_bank_path:
            self.log(f"  ✓ 无个税银行报盘：{os.path.basename(no_tax_bank_path)}")

        self.log("")
        self.log("合并完成，未执行打印。")
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

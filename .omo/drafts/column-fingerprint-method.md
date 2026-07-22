# 列指纹方法 — 处理混合列数变体的工资表合并

## 问题场景

同一组合并组内，各源工资表有不同的列数变体（29/30/31/33列），同一逻辑列（如"单位代理费"）在不同变体中的列号不同。按列号映射会导致数据错位。

## 核心思路

用 **`(行3值, 行4值, 行5值)` 三元组**作为每列的"指纹"，按指纹内容匹配列身份，不依赖列号位置。

## 操作步骤

### 1. 填充合并格

每个源文件的行3-行5是3层复合表头，存在横向合并（如"扣款明细"跨C17:C28）和纵向合并（如"工伤险"跨行4-行5）。

**先把行3-行5范围内的合并格左上角值扩散到整个合并区域**（横竖都填），这样直接读 Cells 值就能得到正确信息。

```python
merges = list(src_ws.merged_cells.ranges)
for mc in merges:
    r1, r2 = mc.min_row, mc.max_row
    c1, c2 = mc.min_col, mc.max_col
    if r2 >= 3 and r1 <= 5:                     # 与表头行有交叠
        tl = src_ws.cell(row=max(r1, 3), column=c1).value
        if tl is not None:
            tl_s = str(tl).strip()
            for rr in range(max(r1, 3), min(r2, 5) + 1):
                for cc in range(c1, c2 + 1):
                    src_ws.cell(row=rr, column=cc).value = tl_s
```

### 2. 生成列指纹

对每列生成指纹 `(r3v, r4v, r5v)`：

```python
for c in range(1, max_cols + 1):
    r3v = src_ws.cell(row=3, column=c).value
    r4v = src_ws.cell(row=4, column=c).value
    r5v = src_ws.cell(row=5, column=c).value
    fp = (str(r3v).strip() if r3v else "",
          str(r4v).strip() if r4v else "",
          str(r5v).strip() if r5v else "")
```

### 3. 建立规范列序

以最普遍的变体（通常是30列）为参考，按参考文件的列顺序建立规范指纹序列，然后追加其他变体独有的指纹：

```python
canonical_fps = []
seen_fp = set()

# 参考文件的列序
for c in range(3, max_cols + 1):
    fp = ref_fp.get(c, ("", "", ""))
    if fp != ("", "", "") and fp not in seen_fp:
        canonical_fps.append(fp)
        seen_fp.add(fp)

# 其他变体的额外列
for fpf in file_fingerprints:
    for c, fp in fpf.items():
        if c >= 3 and fp != ("", "", "") and fp not in seen_fp:
            canonical_fps.append(fp)
            seen_fp.add(fp)
```

### 4. 每个文件的列映射

每个文件建立 `{指纹 → 源列号}` 反向索引，再构建 `{规范位置 → 源列号}`：

```python
rev = {fp: src_c for src_c, fp in fpf.items()}
col_map = {}
for vi, fp in enumerate(canonical_fps):
    if fp in rev:
        col_map[vi] = rev[fp]
```

### 5. 写数据

用每个文件自己的映射取出合计行值：

```python
for vi in range(len(canonical_fps)):
    src_c = col_map.get(vi)
    if src_c is not None and src_c in file_totals:
        val = file_totals[src_c]
        # 写入虚拟表位置 vi+3
```

### 6. 重建合并规则

表头合并从指纹模式推导，不依赖源文件 `merged_cells.ranges`：

| 模式 | 合并方式 | 示例 |
|------|---------|------|
| 连续多列 r3v 相同 | 行3跨列合并 | "扣款明细"跨C17:C28 |
| 行3同组内连续多列 r4v 相同 | 行4跨列合并 | "养老"跨单位/个人2列 |
| r4v有值且r5v为空 | 行4-行5纵向合并 | "工伤险"、"单位代理费" |
| r4v和r5v都为空 | 跨3行合并 | 序号、结算单元名称 |

## 变体对照表

以下为当前已知变体的列含义对照（C7起的数据列）：

| 列号 | 29列 | 30列（标准） | 31列 | 33列 |
|------|------|-------------|------|------|
| C7 | 基本工资 | 基本工资 | 基本工资 | 基本工资 |
| C8 | 交通补贴 | 交通补贴 | 补发工资 | 交通补贴 |
| C9 | 应发工资 | 应发工资 | 交通补贴 | 应发工资 |
| C10 | 单位代理费 | 单位缴纳五险一金 | 应发工资 | 单位缴纳五险一金 |
| C11 | 转账合计 | 单位代理费 | 单位缴纳五险一金 | 单位代理费 |
| C12 | 社保基数 | 转账合计 | 单位代理费 | 大病险 |
| C13 | 医保基数 | 社保基数 | 转账合计 | 转账合计 |
| C14 | 工伤基数 | 医保基数 | 社保基数 | 社保基数 |
| C15 | 公积金基数 | 工伤基数 | 医保基数 | 医保基数 |
| C16 | 扣款明细 | 公积金基数 | 工伤基数 | 工伤基数 |
| C17 | (扣款子列) | 扣款明细 | 公积金基数 | 公积金基数 |
| C18+ | (扣款子列) | (扣款子列) | 扣款明细 | 扣款明细 |

## 优势

1. **兼容性**：新增变体无需修改代码，自动匹配
2. **正确性**：按列名（指纹）匹配，不依赖列号
3. **可溯源性**：每列的含义由源文件自己的表头定义，不依赖约定

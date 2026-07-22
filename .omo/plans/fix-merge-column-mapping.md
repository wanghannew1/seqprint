# 修复虚拟表列错位 + 合计行检测

## TL;DR

> **修复**: `merge_payrolls_simple()` 中不同列数变体（29/30/31/33列）混合在同一组时数据错位的问题，以及合计行检测误匹配的问题。
>
> **方法**: 用列指纹 `(r3v, r4v, r5v)` 替代按列号映射，每个源文件独立检测列含义，数据按规范列序对齐。

---

## 问题分析

### 问题1：不同列数变体混在同一组导致数据错位

吉林大学组 97 个文件中有 4 种列数变体：

| 变体 | 数量 | C7 | C8 | C9 | C10 | C11 | C12 | C13 |
|------|------|----|----|----|-----|-----|-----|-----|
| 29列 | 12 | 基本工资 | 交通补贴 | 应发工资 | 单位代理费 | 转账合计 | 社保基数 | ... |
| 30列 | 73 | 基本工资 | 交通补贴 | 应发工资 | 单位缴纳五险一金 | 单位代理费 | 转账合计 | ... |
| 31列 | 9 | 基本工资 | 补发工资 | 交通补贴 | 应发工资 | 单位缴纳五险一金 | 单位代理费 | ... |
| 33列 | 3 | 基本工资 | 交通补贴 | 应发工资 | 单位缴纳五险一金 | 单位代理费 | 大病险 | ... |

当前代码用统一列号 `active_cols` → 29列文件的 `C10=单位代理费` 填到"单位缴纳五险一金"下面，数据错位。

### 问题2：合计行检测误匹配

`"合计" in str(v)` 会匹配到"每月转账合计8500"、"转账合计15000"等行，28个文件的合计行检测错误，导致虚拟表中对应单位无数据。

---

## 修法思路

核心思想：**用列指纹 `(行3值, 行4值, 行5值)` 替代列号来识别列**。

### 流程图

```
For each 源文件:
  1. 读 merged_cells.ranges（行3-5范围内）
  2. 把合并格左上角值 → 扩散填充到整个合并区域（横竖都填）
  3. 直接读行3-5的 Cells 值 → 每列生成指纹 (r3v, r4v, r5v)
  4. 读合计行 → {源列号 → 值}
  
构建规范列序:
  5. 以最普遍的变体（30列）为参考
  6. 参考文件的指纹顺序作为规范列序
  7. 追加其他变体独有的指纹
  
每个文件的映射:
  8. {指纹 → 源列号} → {规范位置 → 源列号}
  
写虚拟表:
  9. 表头行: 规范指纹序的 (r3v, r4v, r5v) 逐行写入
  10. 合并: 从指纹模式重建（连续相同 r3v→行3合并，连续相同 r4v→行4合并等）
  11. 数据: 用每个文件自己的 {规范位置→源列号} 映射取出合计行值
```

### 关键技术点

**合并格填充**（行3-5范围内，横竖都填）：
```python
for mc in src_ws_r.merged_cells.ranges:
    r1, r2 = mc.min_row, mc.max_row
    c1, c2 = mc.min_col, mc.max_col
    if r2 >= 3 and r1 <= 5:
        tl = src_ws_r.cell(row=max(r1, 3), column=c1).value
        if tl is not None:
            tl_s = str(tl).strip()
            for rr in range(max(r1, 3), min(r2, 5) + 1):
                for cc in range(c1, c2 + 1):
                    src_ws_r.cell(row=rr, column=cc).value = tl_s
```

**列指纹生成**：
```python
for c in range(1, max_cols + 1):
    r3v = src_ws_r.cell(row=3, column=c).value
    r4v = src_ws_r.cell(row=4, column=c).value
    r5v = src_ws_r.cell(row=5, column=c).value
    fp = (str(r3v).strip() if r3v else "",
          str(r4v).strip() if r4v else "",
          str(r5v).strip() if r5v else "")
    fp_cols[c] = fp
```

**规范列序构建**：
```python
# 参考文件（最常见变体）的指纹顺序
canonical_fps = []
seen_fp = set()
for c in range(3, max_cols + 1):
    fp = ref_fp.get(c, ("", "", ""))
    if fp != ("", "", "") and fp not in seen_fp:
        canonical_fps.append(fp)
        seen_fp.add(fp)
# 追加其他变体独有的指纹
for fpf in file_fingerprints:
    for c, fp in fpf.items():
        if c >= 3 and fp != ("", "", "") and fp not in seen_fp:
            canonical_fps.append(fp)
            seen_fp.add(fp)
```

**合并规则重建**（从指纹模式推导）：
- 行3合并：连续相同的 r3v → 跨列合并（如"扣款明细"跨多列）
- 行4合并：在行3同组内，连续相同的 r4v → 跨列合并（如"养老"跨单位/个人2列）
- 纵向合并：r4v有值且r5v为空 → 单列纵向合并行4-行5（如"工伤险"）
- 剩余：行4+行5都为空 → 跨3行合并（序号、单位名称等）

---

## TODOs

- [x] 1. 改造 `merge_payrolls_simple()`：替换 `active_cols` + `hdr_rows` + `src_merges` 为指纹映射机制

  **改动范围**：`batchprint_gui.py` 中 `merge_payrolls_simple()` 函数，从 `# 读取每个源文件的合计行` 到 `r += 2`（约180行）

  **具体变更**：
  1. 合计行检测：`"合计" in str(v)` → `str(v).strip() == "合计"`
  2. 新增合并格填充逻辑（行3-5范围内横竖扩散）
  3. 新增列指纹构建逻辑
  4. 新增规范列序构建逻辑（以最常见变体为参考）
  5. 删除旧 `active_cols` + `hdr_rows` + `src_merges` 逻辑
  6. 重写表头写入：按规范指纹序的 (r3v, r4v, r5v) 逐行写入
  7. 重写合并逻辑：从指纹模式重建合并规则
  8. 重写数据写入：用 `{规范位置→源列号}` 映射填合计行值
  9. 标题行和单位名称行移到表头写入之后（因为 `virtual_cols` 现在由规范指纹序决定）

  **不改变**：WPS 复制源表部分（第694行之后）不变

  **验证**：用白云数据跑一遍，检查虚拟表数据对齐



---

## Final Verification Wave

- [ ] F1. **验证：用白云目录跑一次** — `quick`
  用白云源文件目录跑合并，检查：
  - 29列 vs 30列 vs 31列 vs 33列 文件的数据是否正确填到对应列下
  - 之前 28 个合计行检测错误的单位现在有数据了
  - 合并单元格视觉上正确（扣款明细跨列、养老/失业各跨2列等）

---

## 文件路径
- `batchprint_gui.py`（`merge_payrolls_simple()` 函数）

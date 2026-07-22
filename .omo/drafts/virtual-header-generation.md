# 虚拟工资集合表复合表头：读取 → 写入 → 合并

> 对应代码：`batchprint_gui.py` → `merge_payrolls_simple()` 函数

---

## 一、总体布局

虚拟表从上到下的结构：

```
行1: 大标题        ← {单位名} 工资表合集（{年月}），整行水平合并
行2: 单位信息      ← 左"单位名称：{单位名}" + 右"统计时间：{日期}"
行3: 三级表头-第1层 ← 序号 | 结算单元名称 | [规范指纹序的第3行值]
行4: 三级表头-第2层 ← (空) | (空)        | [规范指纹序的第4行值]
行5: 三级表头-第3层 ← (空) | (空)        | [规范指纹序的第5行值]
行6+: 每个源文件一条合计行
  合计行: 序号 | 单位名 | [各列合计值]
行N+2: (空行)
行N+3+: 源文件原样堆叠（含签名图片，通过 WPS COM Paste 复制）
```

表头列数 = 2（序号 + 结算单元名称）+ `len(canonical_fps)`（工资科目列）。

---

## 二、关键数据结构

### 指纹 `fp = (r3v, r4v, r5v)`

每列的唯一标识，由该列行3/行4/行5的值组成。例如：
- `("基本工资", "基本工资", "基本工资")` — 单层列头，3行值相同
- `("扣款明细", "养老", "单位")` — 三层复合列头

指纹用在两个地方：
1. **数据映射**：作为字典 key，将各源文件的列映射到虚拟表统一位置
2. **显示回退**：变体独有列不在 `display_hdr` 中时，回退到指纹自身的值作为显示文字

### 参考文件

以最常见的列数变体为参考（白云97个文件中30列占73个，选30列）。参考文件提供：
- 指纹序列顺序（`canonical_fps`）
- 显示表头的原始文字（`display_hdr`）

---

## 三、流程概览

```
每个源文件:
  └─ 合并格展开 → 提取 fp_dict: {列号 → (r3v, r4v, r5v)}
  └─ 读取合计行

选参考文件:
  └─ 统计各变体列数 → 选列数最多的为参考

构建规范列序 canonical_fps:
  └─ 参考文件 C7+ 指纹 → 去重
  └─ 其他变体独有指纹 → 位置感知插入（同r3v组内，合计前）

过滤空列:
  └─ 所有文件该列合计值全为0/空 → 从 canonical_fps 移除

读取参考文件表头:
  └─ 合并格展开 → ref_hdr[3行][max_cols+1]

构建 display_hdr:
  └─ 每个参考文件指纹 → (ref_hdr行3值, 行4值, 行5值)

写入表头:
  └─ 行1: 大标题
  └─ 行2: 单位名 + 统计时间
  └─ 行3-5: 遍历 canonical_fps → display_hdr.get(fp, fp)[hi]

合并单元格:
  └─ [1] 行4水平合并（按行3分组，连续相同值）
  └─ [2] 纵向合并（情形A/B/C）
  └─ [3] 兜底3行合并（未合并且r4+r5为空）
  └─ [4] 行3水平合并（最后做，避免COM读已合并格返回Empty）
```

---

## 四、指纹提取（每个源文件独立执行）

```python
# 构建合并格查找表
# 行3-5范围内的合并区域，非左上角格填左上角值
merge_lookup = {}
for mc in src_ws.merged_cells.ranges:
    if mc.max_row >= 3 and mc.min_row <= 5:
        tl = src_ws.cell(row=mc.min_row, column=mc.min_col).value
        if tl is not None:
            tl_s = str(tl).strip()
            for rr in range(max(mc.min_row, 3), min(mc.max_row, 5) + 1):
                for cc in range(mc.min_col, mc.max_col + 1):
                    merge_lookup[(rr, cc)] = tl_s

def _cv(row, col):
    """取单元格值，合并格返回左上角值"""
    if (row, col) in merge_lookup:
        return merge_lookup[(row, col)]
    v = src_ws.cell(row=row, column=col).value
    return str(v).strip() if v is not None else ""

# 生成列指纹
for c in range(1, src_ncols + 1):
    fp = (_cv(3, c), _cv(4, c), _cv(5, c))
    fp_dict[c] = fp
```

> `merge_lookup` 只在行3-5范围内生效，跳过人员信息行（行1-2）和工资表数据区域（行6+）。

---

## 五、构建规范列指纹序列 `canonical_fps`

### 5.1 选参考文件

统计每个源文件 C3+ 的非空指纹数量，以出现最多的列数为参考：

```python
variant_counts = [sum(1 for c, fp in fpf.items() if c >= 3 and fp != ("", "", "")) for fpf in file_fingerprints]
count_freq = Counter(variant_counts)
most_common_cnt = max(count_freq, key=count_freq.get)
ref_idx = next(idx for idx, c in enumerate(variant_counts) if c == most_common_cnt)
ref_fp = file_fingerprints[ref_idx]
```

### 5.2 参考文件指纹序列（C7起，去重）

C1-C6 是人员信息列（序号/姓名/身份证/部门/岗位/职工号），所有文件固定一致，不参与指纹映射。虚拟表前2列固定为"序号"和"结算单元名称"，直接覆盖 C1-C6。

```python
canonical_fps = []
for c in range(7, max_cols + 1):
    fp = ref_fp.get(c, ("", "", ""))
    if fp != ("", "", "") and fp not in canonical_fps:
        canonical_fps.append(fp)
```

### 5.3 变体独有指纹：位置感知插入

其他变体文件可能包含参考文件没有的列（如"公务员医疗补助"存在于31列变体，但30列参考中没有）。追加到末尾会导致排序错误——"公务员医疗补助"出现在"扣款合计"后面。

插入规则：
1. 在同 r3v 组中找第一个"合计"项 → 插到它前面
2. 若无"合计"项 → 插到同 r3v 组末尾（其他组之前）
3. 若无可匹配的 r3v 组 → 追加到尾部

```python
for fpf in file_fingerprints:
    for c, fp in fpf.items():
        if c >= 7 and fp != ("", "", "") and fp not in seen_fp:
            seen_fp.add(fp)
            insert_idx = len(canonical_fps)  # 默认追加
            same_r3v_last = -1
            for idx, existing_fp in enumerate(canonical_fps):
                if existing_fp[0] == fp[0]:          # 同 r3v 组
                    same_r3v_last = idx
                    if "合计" in existing_fp[1]:     # 本组有合计项
                        insert_idx = idx
                        break
            if insert_idx == len(canonical_fps) and same_r3v_last >= 0:
                insert_idx = same_r3v_last + 1       # 插到本组末尾
            canonical_fps.insert(insert_idx, fp)
```

### 5.4 过滤空列

遍历每个指纹列在所有源文件合计行中的数据。若全部为0或空值，从 `canonical_fps` 中移除（虚拟表中不展示）：

```python
active_fps = []
for vi, fp in enumerate(canonical_fps):
    has_data = False
    for ft, cm in zip(file_totals, file_col_maps):
        src_c = cm.get(vi)
        if src_c is not None and src_c in ft:
            v = ft[src_c]
            if v is not None:
                try:
                    if float(v) != 0:
                        has_data = True; break
                except:
                    if str(v).strip():
                        has_data = True; break
    if has_data:
        active_fps.append(fp)
canonical_fps = active_fps or canonical_fps[:1]  # 至少保留1列
```

---

## 六、生成显示表头 `display_hdr`

`canonical_fps` 中的指纹用于数据映射，但显示时需要原始表头文字。通过读取参考文件的行3-5来构建。

### 6.1 读取参考文件 + 合并格展开

openpyxl 读合并格的非左上角返回 None。用 `merged_cells.ranges` 展开：

```python
from openpyxl.utils import range_boundaries

ref_hdr = [[""] * (max_cols + 1) for _ in range(3)]

# Step 1: 读原始值
for hi, hr in enumerate([3, 4, 5]):
    for c in range(1, max_cols + 1):
        v = ref_ws.cell(row=hr, column=c).value
        if v is not None:
            ref_hdr[hi][c] = str(v).strip()

# Step 2: 合并格展开——非左上角填左上角值
for mr in ref_ws.merged_cells.ranges:
    mc, mr0, Mc, Mr = range_boundaries(str(mr))
    if Mr < 3 or mr0 > 5:
        continue  # 不在表头范围内
    for hi, hr in enumerate(range(max(mr0, 3), min(Mr, 5) + 1)):
        tl = ref_hdr[hr - 3][mc]
        if not tl:
            continue
        for c in range(mc, Mc + 1):
            ref_hdr[hr - 3][c] = tl
```

展开前后的对比示例：

```
展开前：   C20                C21
行3:      "扣款合计"          "个人所得税"
行4:      "扣款合计"          ""          ← openpyxl 返回 None
行5:      "扣款合计"          ""          ← openpyxl 返回 None

展开后：
行3:      "扣款合计"          "个人所得税"
行4:      "扣款合计"          "个人所得税"  ← 合并格展开填上
行5:      "扣款合计"          "个人所得税"  ← 合并格展开填上
```

> 注：指纹提取（第四节）和表头读取（本节）都用到合并格展开，但写法不同。指纹提取用 `merge_lookup` dict，表头读取用 `range_boundaries` 直接操作列表——两者等价。

### 6.2 构建映射

```python
display_hdr = {}
for c in range(7, max_cols + 1):
    fp = ref_fp.get(c)
    if fp and fp != ("", "", ""):
        display_hdr[fp] = (ref_hdr[0][c], ref_hdr[1][c], ref_hdr[2][c])
```

**变体独有列的回退**：不在参考文件中的指纹（如公务员医疗补助）在 `display_hdr` 中无记录。写表头时通过 `display_hdr.get(fp, fp)` 回退到指纹元组自身的值——指纹中的 r4v/r5v 就是源文件中的列头文字，正确可用。

---

## 七、写入表头

### 7.1 行1：大标题

```python
title = f"{group_key} 工资表合集（{ym_display}）"
tgt_ws.Cells(r, 1).Value = title
tgt_ws.Range(tgt_ws.Cells(r, 1), tgt_ws.Cells(r, virtual_cols)).Merge()
tgt_ws.Range(...).Font.Bold = True
tgt_ws.Range(...).Font.Size = 16
tgt_ws.Range(...).HorizontalAlignment = -4108  # xlHAlignCenter
```

### 7.2 行2：单位信息（左）+ 统计时间（右）

左半段从C1开始合并，右半段从 `right_start` 到末尾合并，互不重叠：

```python
right_start = virtual_cols - 3 if virtual_cols > 6 else 4
tgt_ws.Cells(r, 1).Value = f"单位名称：{group_key}"
tgt_ws.Range(tgt_ws.Cells(r, 1), tgt_ws.Cells(r, right_start - 1)).Merge()
tgt_ws.Range(...).HorizontalAlignment = 1  # xlHAlignLeft
tgt_ws.Cells(r, right_start).Value = f"统计时间：{datetime.now().strftime('%Y年%m月%d日')}"
tgt_ws.Range(tgt_ws.Cells(r, right_start), tgt_ws.Cells(r, virtual_cols)).Merge()
tgt_ws.Range(...).HorizontalAlignment = -4152  # xlHAlignRight
```

### 7.3 行3-5：三级复合表头

```python
hdr_start_row = r
for hi in range(3):          # hi=0:行3, hi=1:行4, hi=2:行5
    tgt_ws.Cells(r, 1).Value = "序号" if hi == 0 else ""
    tgt_ws.Cells(r, 2).Value = "结算单元名称" if hi == 0 else ""
    for vi, fp in enumerate(canonical_fps, 3):
        dv = display_hdr.get(fp, fp)[hi]
        tgt_ws.Cells(r, vi).Value = dv
        if hi == 0 and dv:
            tgt_ws.Cells(r, vi).Font.Bold = True
    r += 1
hdr_end_row = r - 1

# 居中对齐
for hr in range(hdr_start_row, hdr_end_row + 1):
    tgt_ws.Range(tgt_ws.Cells(hr, 1), tgt_ws.Cells(hr, virtual_cols)).HorizontalAlignment = -4108
```

写入后的裸数据（30列参考 + 31列变体含公务员医疗补助）：

```
行3: 序号 | 结算单元名称 | 基本工资 | 交通补贴 | 应发工资 | ... | 扣款明细 | ... | 个人所得税 | 实发工资 | 实发合计
行4:      |              | 基本工资 | 交通补贴 | 应发工资 | ... | 养老     | ... | 个人所得税 | 实发工资 | 实发合计
行5:      |              | 基本工资 | 交通补贴 | 应发工资 | ... | 单位     | ... | 个人所得税 | 实发工资 | 实发合计
```

此时各单元格独立，尚未合并。

---

## 八、合并单元格

扫描已写入的内容，相邻且值相同的合并。**不依赖源文件的 `merged_cells.ranges`**。

### 8.1 合并顺序（关键约束）

**行4合并 → 纵向合并 → 兜底3行合并 → 行3合并**

行3合并在最后。因为 WPS COM 读取已合并格的非左上角返回 Empty，若先行3合并再读行4，"扣款明细"合并范围内的大量列行3变成 Empty → 行4按行3分组时每列独立 → 养老/失业等子项无法合并。

### 8.2 防冲突机制

`hmerged = set()` 记录已合并的单元格坐标 `(行号, 列号)`。每步合并前检查目标范围是否已在 hmerged 中，避免重复合并引发 COM 异常。

### 8.3 行4水平合并

在行3分组内（行3尚未合并，值可读），合并行4连续相同值：

```python
vi = 3
while vi <= virtual_cols:
    r3v = tgt_ws.Cells(hdr_start_row, vi).Value
    # 找行3同组的范围 [vi, vj)
    vj = vi + 1
    while vj <= virtual_cols and tgt_ws.Cells(hdr_start_row, vj).Value == r3v:
        vj += 1
    # 在行3分组内，合并行4连续相同值
    vk = vi
    while vk < vj:
        r4v = tgt_ws.Cells(hdr_start_row + 1, vk).Value
        if not r4v:
            vk += 1
            continue
        vl = vk + 1
        while vl < vj and tgt_ws.Cells(hdr_start_row + 1, vl).Value == r4v:
            vl += 1
        if vl - 1 > vk:
            tgt_ws.Range(tgt_ws.Cells(hdr_start_row + 1, vk),
                         tgt_ws.Cells(hdr_start_row + 1, vl - 1)).Merge()
        vk = vl
    vi = vj
```

```
行3(未合并): [扣款明细] [扣款明细] [扣款明细] [扣款明细] [扣款明细]
行4(合并前): [养老]     [养老]     [失业]     [公务员医疗补助] [单位代理费]
行4(合并后): [──养老──] [失业]     [公务员医疗补助] [单位代理费]
```

### 8.4 纵向合并

三种情形：

| 情形 | 条件 | 例子 | 合并范围 |
|------|------|------|---------|
| A | r3v==r4v==r5v | 基本工资、个人所得税、实发合计 | 行3-5（3行全合） |
| B | 仅r4v==r5v, r3v不同 | 单位代理费、公务员医疗补助 | 行4-5（2行合） |
| C | r4v有值, r5v为空 | 工伤险 | 行4-5（2行合） |

```python
for vi in range(3, virtual_cols + 1):
    r3v = tgt_ws.Cells(hdr_start_row, vi).Value
    r4v = tgt_ws.Cells(hdr_start_row + 1, vi).Value
    r5v = tgt_ws.Cells(hdr_start_row + 2, vi).Value
    if r4v and (not r5v or r4v == r5v):
        if r3v and r3v == r4v:
            # 情形A：合3行
            tgt_ws.Range(tgt_ws.Cells(hdr_start_row, vi),
                         tgt_ws.Cells(hdr_end_row, vi)).Merge()
        else:
            # 情形B/C：合2行（行4-5）
            tgt_ws.Range(tgt_ws.Cells(hdr_start_row + 1, vi),
                         tgt_ws.Cells(hdr_end_row, vi)).Merge()
```

排除项：r4v != r5v 的列不参与纵向合并（如"养老/单位"和"养老/个人"是两列不同子项）。

### 8.5 兜底3行合并

未参与任何合并的列，若行4+行5都为空，合并行3-5（如"序号"列）：

```python
for vi in range(1, virtual_cols + 1):
    if any((rr, vi) in hmerged for rr in range(hdr_start_row, hdr_end_row + 1)):
        continue
    r4v = tgt_ws.Cells(hdr_start_row + 1, vi).Value
    r5v = tgt_ws.Cells(hdr_start_row + 2, vi).Value
    if not r4v and not r5v:
        tgt_ws.Range(tgt_ws.Cells(hdr_start_row, vi),
                     tgt_ws.Cells(hdr_end_row, vi)).Merge()
```

### 8.6 行3水平合并（最后执行）

```python
vi = 3
while vi <= virtual_cols:
    r3v = tgt_ws.Cells(hdr_start_row, vi).Value
    if not r3v:
        vi += 1
        continue
    vj = vi + 1
    while vj <= virtual_cols and tgt_ws.Cells(hdr_start_row, vj).Value == r3v:
        vj += 1
    if vj - 1 > vi:
        tgt_ws.Range(tgt_ws.Cells(hdr_start_row, vi),
                     tgt_ws.Cells(hdr_start_row, vj - 1)).Merge()
    vi = vj
```

```
合并前: [扣款明细] [扣款明细] [扣款明细] [扣款明细] [扣款明细]
合并后: [──────────────── 扣款明细 ────────────────]
```

### 8.7 合并效果总览

以一笔"扣款明细"区域为例（行4合并 + 行3合并 + 纵向合并）：

```
       C9     C10    C11    C12    C13    C14    C15    C16    C17    C18    C19    C20
行3:   [────────────────────────────── 扣款明细 ──────────────────────────────────────]
行4:   [─── 养老 ───] [── 失业 ──] [公务员医疗补助] [单位代理费] [─── 扣款合计 ─────]
行5:     单位    个人    单位    个人  公务员医疗补助  单位代理费  扣款合计
       └ 纵向合并 ┘                        └ 纵向合并 ┘  └ 纵向合并 ┘ └ 纵向合并  ┘
       C9-C12无纵向合并（r4v != r5v）
```

---

## 九、数据写入

每个源文件的数据通过 `rev = {fp: 源列号}` 反向映射写入虚拟表对应列。详见主文档《fingerprint-method.md》。

---

## 十、代码地图

| 步骤 | 代码位置 | 说明 |
|------|---------|------|
| 提取各文件指纹 | 线 494-539 | 合并格展开 + `fp_dict: 列号→(r3v,r4v,r5v)` |
| 选参考文件 | 线 543-551 | 统计列数出现频次，取最常见变体 |
| 构建 canonical_fps | 线 556-580 | 参考指纹去重 + 变体独有位置感知插入 |
| 列映射 | 线 583-591 | 各文件 `rev[fp]` 定位源列号 |
| 过滤空列 | 线 593-616 | 全零合计列移除 |
| 重新映射 | 线 618-626 | 过滤后位置重算 |
| 读取参考表头 | 线 628-654 | `merged_cells.ranges` 展开 |
| 构建 display_hdr | 线 656-665 | 指纹→显示值 |
| 写行1-2 | 线 667-688 | 大标题 + 单位名/统计时间 |
| 写行3-5 | 线 690-704 | 遍历 canonical_fps 写3级表头 |
| 居中对齐 | 线 707-709 | 全表头范围 |
| 行4水平合并 | 线 716-741 | 行3分组内连续相同值合并 |
| 纵向合并 | 线 743-771 | 情形A/B/C |
| 兜底3行合并 | 线 773-784 | 未合并且r4+r5为空 |
| 行3水平合并 | 线 786-802 | 最后执行（COM独眼龙问题） |

---

## 十一、常见问题

### Q: 为什么不用复制源文件的 merged_cells.ranges？

旧方案就是复制合并格，但：
- 不同变体列数不同，合并格位置不同
- 变体独有列的合并格在参考文件中不存在
- 空列过滤后列映射不连续，合并格无法正确映射

指纹方法不依赖具体的合并格位置，通过内容匹配和写入值扫描动态合并，自适应所有变体。

### Q: 公务员医疗补助为什么会在扣款合计后面？

变体独有指纹追加到 `canonical_fps` 末尾时，落在"扣款合计"后面。而各组 r3v 在规范序列中非连续（被个人所得税/实发工资等穿插），组内重排无法跨段修复。修复方式：位置感知插入——变体指纹插入到同 r3v 组第一个"合计"项之前。

### Q: 行4子项合并（养老/失业等）为什么曾经不生效？

**第一层（display_hdr 空值）**：openpyxl 读合并格非左上角返回 None，`ref_hdr[1][c]` 为空 → `display_hdr` 行4值为空 → 写表头时空白。最终修复：用 `merged_cells.ranges` 展开（提交 `203de90`）。

**第二层（合并顺序错误）**：行4合并先于行3合并——COM 读已合并格非左上角返回 Empty，导致行3合并后行4按行3分组时被跳过。修复：交换顺序，行4合并在行3之前（提交 `311ff85`）。

### Q: 行5为什么没有纵向传播？

行5的子项都是独立值（"单位"、"个人"），无跨列合并。传播行5会污染后续列。

# 打印调整方案

## 目标

工资表通过 WPS 打印时：
1. 所有列在一页内（自动缩放到一页宽）
2. 表头（第 1-5 行）每页重复
3. 同一行不跨页截断
4. 有边框、合并单元格、适当行高列宽

---

## 当前 `write_payroll` 中的设置（openpyxl）

```python
ws.page_setup.orientation = 'landscape'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.page_margins.left = 0.5
ws.page_margins.right = 0.3
ws.page_margins.top = 0.5
ws.page_margins.bottom = 0.5
```

写在 `batchprint_gui.py` 的 `merge_payrolls_by_tax()` → `write_payroll()` 中。

---

## 缺失的设置（需补充）

### 1. 打印标题行（每页重复表头）

```python
# openpyxl 方式
ws.print_title_rows = '1:5'
```

或者用 COM API 方式（适合后续 WPS 打开时生效）：

```python
ws.api.PageSetup.PrintTitleRows = "$1:$5"
```

作用：第 1-5 行（标题 + 单位 + 三行表头）在每一页顶部重复打印。

### 2. Zoom = False（配合 FitToPages）

openpyxl 的 `fitToWidth`/`fitToHeight` 会自动设 Zoom 为 False。当前已设：
```python
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
```

效果等价于 COM 方式：
```python
ws.api.PageSetup.Zoom = False
ws.api.PageSetup.FitToPagesWide = 1
ws.api.PageSetup.FitToPagesTall = False
```

### 3. 页边距（参考 Legacy）

源文件（Legacy）使用厘米单位：
```python
CM_TO_POINTS = 28.35
ws.api.PageSetup.LeftMargin = 2 * CM_TO_POINTS  # 2cm
```

openpyxl 的 `page_margins` 以英寸为单位：
| 英寸 | 厘米 |
|------|------|
| 0.5 | ~1.27 |
| 0.75 | ~1.90 |
| 1.0 | ~2.54 |

建议：左/右边距 0.5 英寸（折中），或改为 0.3 以腾出更多列空间。

---

## Legacy 参考代码

位置：`/home/ubuntu/coding/legacy/XForm_tkinterdnd2_optimized.py` 第 288-292 行

```python
CM_TO_POINTS = 28.35

ws.api.PageSetup.LeftMargin = 2 * CM_TO_POINTS    # 左边距 2cm
ws.api.PageSetup.PrintTitleRows = "$1:$5"           # 表头重复行
ws.api.PageSetup.Zoom = False                       # 关闭缩放比例
ws.api.PageSetup.FitToPagesWide = 1                 # 所有列适应一页宽
ws.api.PageSetup.FitToPagesTall = False             # 行高自适应
```

注意：`ws.api` 是 openpyxl 的 COM 兼容接口，直接传递参数给底层引擎。写入 xlsx 后 WPS 打开时也生效。

---

## 当前效果对比

| 项目 | 之前（无设置） | 现在（openpyxl） | 目标（Legacy 方式） |
|------|---------------|-------------------|---------------------|
| 列缩放 | 默认 | fitToWidth=1 | FitToPagesWide=1 |
| 行缩放 | 默认 | fitToHeight=0 | FitToPagesTall=False |
| 表头重复 | — | — | PrintTitleRows="$1:$5" |
| 边框 | 无 | thin | thin |
| 表头合并 | 仅 2 处 | 12 处（自动相邻合并）| 按需 |
| 行高 | 默认 | 36/28/28/28/28 | 需测试 |
| 列宽 | 默认 | 6~20 | 需测试 |

---

## 实施步骤

1. 在 `write_payroll()` 中补充：
   ```python
   ws.print_title_rows = '1:5'
   ```
2. 如需更精确控制页边距（厘米单位），用：
   ```python
   ws.page_margins.left = 0.5   # ~1.27cm
   ws.page_margins.right = 0.3  # ~0.76cm
   ```
3. 验证生成的 xlsx 在 WPS 打印预览中：
   - 所有列显示在一页内
   - 第 1-5 行重复
   - 行不跨页截断

---

## 注意事项

- `fitToHeight = 0` 表示"自动"，行数可能跨多页，但不截断单行
- 如需强制所有行在一页：设 `fitToHeight = 1`（但 200+ 行字体会极小）
- 列宽不宜过宽（扣款明细子列 9 字符足够），否则缩放比例太小
- 数据行使用 `center_align_nowrap`（不换行），避免行高膨胀

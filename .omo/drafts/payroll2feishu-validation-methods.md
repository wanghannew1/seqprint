# submit_approval_feishu — 表格数据关系验证方法总结

> 项目：Payroll2Feishu（工资单推送飞书OA审批）
> 核心文件：`demo_app.py`（~1827 行 Streamlit 单文件应用）
> 配置文件：`config.json`（验证规则可配置化）

---

## 一、验证架构总览

系统采用 **配置驱动 + 代码引擎** 分层架构：

```
config.json validation 段                              demo_app.py
┌──────────────────────────────┐         ┌──────────────────────────────────────┐
│ enabled / strict / tolerance  │ ─────→  │ validate_payroll() 核心校验函数      │
│ column_sum_checks  列加总     │         │   ├─ A. 纵向列加总校验               │
│ row_formulas        横向公式  │         │   ├─ B. 合计行横向公式              │
│ required_signatures 签名栏   │         │   ├─ C. 每行横向公式                │
│ write_back_sheet    回写     │         │   └─ D. 表格格式校验                │
└──────────────────────────────┘         ├─ check_signatures() 签名扫描        │
                                         ├─ append_validation_sheet() 回写Excel│
 外部调用                                ├─ build_summary_workbook() 汇总表     │
 main() →                                 └─ _extract_year_month() 年月一致性   │
   标题vs文件名年月一致性检查               └─ 零金额拦截（main 内联逻辑）          │
   零金额阻塞防护                           └─ 全零列自动隐藏（UI 层面）           │
```

- **校验器输出**（`checks` 列表）→ 驱动 **3 条消费链路**：
  1. **Streamlit UI** 实时展示校验结果（通过/失败数目）
  2. **`append_validation_sheet()`** 写入原 Excel 追加 sheet（.xlsx 专用）
  3. **`build_summary_workbook()`** 汇总表中「验证明细」sheet 集中展示

---

## 二、4 种核心校验方法

### A. 纵向列加总校验（column_sum_checks）

**原理**：对配置的某一列，累加所有数据行的值，与合计行该列的值进行比较。

```python
# 伪代码逻辑
col_sum = sum(数据行[列] for 所有数据行)
summary_val = 合计行[列]
diff = col_sum - summary_val  # 精确到分
passed = abs(diff) <= tolerance
```

**配置方式**（`config.json`）：
```json
"column_sum_checks": [
  {"column": "deduction_total"},
  {"column": "personal_tax"},
  {"column": "net_total"}
]
```

**适用场景**：
- 验证明细数据汇总是否等于合计行
- 发现表头配置错误（列映射偏移）导致数据全为 0

**代码位置**：`demo_app.py:796-815`，函数 `validate_payroll()` 段「A」

---

### B. 合计行横向公式校验（row_formula_summary）

**原理**：验证合计行上，等式左侧（LHS）是否等于右侧加减表达式的结果。

```python
# 伪代码逻辑
L = 合计行[lhs列]
R = sum(合计行[rhs_plus列]) - sum(合计行[rhs_minus列])
diff = L - R
passed = abs(diff) <= tolerance
```

**配置方式**：
```json
"row_formulas": [
  {
    "name": "转账合计 = 扣款合计 + ... + 实发合计",
    "lhs": "transfer_total",
    "rhs_plus": ["deduction_total", "personal_tax", "net_total", ...],
    "rhs_minus": ["pay_cash"]
  }
]
```

**适用场景**：
- 验证工资表核心恒等式：**转款合计 = 扣款合计 + 个税 + 各项调整 + 实发合计**
- 发现数据逻辑错误（如某项多录/漏录）

**代码位置**：`demo_app.py:838-851`，函数 `validate_payroll()` 段「B」

---

### C. 数据行横向公式校验（row_formula_rows）

**原理**：对**每一行明细数据**，逐一验证同一恒等式是否成立。

```python
# 伪代码逻辑
for each 数据行:
    L_r = 行[lhs列]
    R_r = sum(行[rhs_plus列]) - sum(行[rhs_minus列])
    if abs(L_r - R_r) > tolerance:
        failed_rows.append(行)
```

**输出**：报告通过行数/总行数 + 前 3 条失败样本明细。

**适用场景**：
- 精确定位具体哪一行数据有误
- 发现个别员工数据录入错误而非整体配置问题

**代码位置**：`demo_app.py:854-885`，函数 `validate_payroll()` 段「C」

---

### D. 表格格式校验（table_format）

**原理**：双重防御性检查，捕获配置与文件不匹配的「静默失败」。

| 子检查 | 条件 | 失败后果 |
|--------|------|---------|
| ① 关键列缺失 | `transfer_total`/`deduction_total`/`net_total` 任一未在表头中找到 | 直接失败，提示调整 `header_start_row` |
| ② 合计行全零 | 三个关键列合计值均为 0，但数据行非空 | 失败，提示格式可能不匹配 |

**设计意图**：当 `header_start_row` 配置与实际文件不符时，列关键字可能通过宽范围搜索命中错误列，导致合计行取值为 0 但校验通过（0=0）。此检查专门捕获这种"安静错误"。

**代码位置**：`demo_app.py:887-936`，函数 `validate_payroll()` 段「D」

---

## 三、辅助验证方法

### E. 签名栏检查（check_signatures）

**原理**：全局扫描 Excel 所有单元格，检查是否包含必要的签名关键词。

**配置方式**（支持分组逻辑）：
```json
"required_signatures": [
  ["总经理签字", "部长签字|部长、分管副总签字|分管领导审核", "财务审核", "业务审核"],
  ["总经理签字", "部长签字|部长、分管副总签字|分管领导审核", "财务审核"]
]
```

- **扁平列表**：所有关键词都必须存在
- **分组列表**：任意一组全部匹配即通过（OR 逻辑）
- **同组支持 `|` 别名**：任一别名匹配即算该组通过

**匹配算法**：大小写不敏感的子串包含匹配（`in` 操作符）。

**代码位置**：`demo_app.py:948-994`

---

### F. 年月一致性检查

**原理**：分别从**报表标题行**和**文件名**提取「YYYY年MM月」，比较两者是否一致。

- 从标题提取：优先权威来源（审计凭证）
- 从文件名提取：作为兜底 + 交叉验证
- 不一致时 UI 黄色警告，但**以标题为准**不阻塞提交

**日期提取算法**（`_extract_year_month`）：
1. 先尝试 4 位年连体：`2026年05月`
2. 再尝试 2 位年连体：`26年5月`（带 50 年消歧规则）
3. 最后尝试分离形态：先找最后一个「N年」，再在其后找「N月」

**代码位置**：`demo_app.py:363-422`（提取函数），`1381-1393`（调用检查）

---

### G. 零金额静默防护

**原理**：提交前检查每个附件的 `transfer_total` 和 `net_total` 是否同时为 0。

**触发条件**：`transfer_total == "0.00"` 且 `net_total == "0.00"`

**阻断逻辑**：禁用提交按钮（`final_blocked = True`），提示用户检查表头配置。

**代码位置**：`demo_app.py:1526-1541`

---

### H. 全零列自动隐藏（UI 层面）

**原理**：在 Streamlit 数据预览中，自动隐藏所有行该列值均为 `0.00` 的列，减少横向滚动。

**不影响**：飞书 API 提交的数据（`parsed_list` / `tf_columns` 不变）。

**代码位置**：`demo_app.py:1460-1470`

---

### I. 数据行汇总兜底

**原理**：当合计行某列为空/None 时，自动从数据行累计汇总该列作为兜底值，而不是直接返回 0。

```python
def get_val(idx, col_key=None):
    if 合计行有值:
        return 该值
    else:
        return _sum_data_rows(col_key)  # 用数据行汇总兜底
```

**代码位置**：`demo_app.py:649-676`

---

## 四、校验结果的两条回写链路

### 链路 1：回写原 Excel（.xlsx 专用）

函数 `append_validation_sheet()` 将校验结果以新 sheet 形式追加到原始 Excel：

```
Sheet: 验证结果
├── 标题：工资表校验结果
├── 元数据：生成时间、源文件、汇总（N 项通过 / M 项失败）
├── 明细表：校验项 | 类型 | 结果(✅/❌) | 说明
│   ├── 纵向加总 / 横向公式 / 表格格式
│   └── 通过行绿色底色，失败行红色底色
└── 列宽自适应
```

**条件**：仅 `.xlsx` 格式（`.xls` 只读不可回写）。

**代码位置**：`demo_app.py:997-1076`

### 链路 2：汇总表集中展示

函数 `build_summary_workbook()` 生成独立汇总表，包含两个 sheet：

| Sheet | 内容 |
|-------|------|
| 汇总数据 | 每个附件一行，含所有审批字段 + 验证结果状态 |
| 验证明细 | 所有附件的校验结果合并在一个表，多一列「附件名」 |

**解决两个问题**：
- `.xls` 无法回写验证结果 → 汇总表统一展示
- 多附件场景无全局视图 → 汇总表集中呈现

**代码位置**：`demo_app.py:1079-1282`

---

## 五、配置驱动的验证框架

### config.json validation 段

```json
{
  "validation": {
    "enabled": true,           // 总开关
    "strict": true,            // true=红色报错禁止提交；false=黄色警告可提交
    "tolerance": 0.00,         // 容差（元），精确到分
    "write_back_sheet": true,  // 是否回写Excel
    "write_back_sheet_name": "验证结果",  // 回写sheet名称
    "column_sum_checks": [     // 列加总校验列表
      {"column": "deduction_total"}
    ],
    "row_formulas": [          // 横向公式校验列表
      {
        "name": "...",
        "lhs": "transfer_total",
        "rhs_plus": ["..."],
        "rhs_minus": ["..."]
      }
    ],
    "required_signatures": [   // 签名栏检查
      ["总经理签字", "部长签字", "财务审核"]
    ]
  }
}
```

### 设计亮点

1. **列名与验证配置解耦**：`excel.columns` 定义列映射（keywords + label），`validation` 通过列 key 引用，修改 Excel 模板的列名无需改代码
2. **缺列零容忍但优雅降级**：`require_col()` 验证列是否在 `excel.columns` 中定义，未定义则抛显式异常；但列已定义却未在表中找到则执行 0 值校验（通过失败结果暴露问题）
3. **strict 模式分离**：严格模式下任何失败阻塞提交；宽松模式仅黄色警告，不阻断业务流程
4. **tolerance 容差**：所有金额精确到分，允许 `±tolerance` 范围内的舍入误差

---

## 六、校验流程总图

```
用户上传 Excel
      │
      ▼
parse_excel()
  ├─ 提取报表标题、单位名、年月
  ├─ 定位合计行（"合计"标记扫描）
  ├─ 定位列位置（配置表头范围 → 宽范围兜底）
  ├─ 提取合计行各列数值 + 数据行列表
  └─ 计算 tax_and_others = transfer - deduction - net
      │
      ▼
validate_payroll()
  ├─ A. column_sum:   sum(数据行) == 合计行
  ├─ B. formula_sum:  合计行 lhs == rhs
  ├─ C. formula_rows: 每行 lhs == rhs（记录失败行）
  └─ D. format:       关键列存在 && 合计行非零
      │
      ▼
check_signatures() ── 扫描全表找签名关键词
      │
      ▼
年月一致性检查 ── 标题年月 vs 文件名年月
零金额防护 ── transfer 和 net 同时为 0？
      │
      ▼
输出：
├─ UI 数据预览（全零列自动隐藏）
├─ 原 Excel 追加「验证结果」sheet（.xlsx）
├─ 汇总表「验证明细」sheet（所有附件合并）
└─ 提交按钮状态（strict 模式下失败则禁用）
```

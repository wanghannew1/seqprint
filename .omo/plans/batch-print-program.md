# 批量打印程序 (batch-print-program)

## TL;DR

> **Quick Summary**: 一个带 GUI 的 Python 工具，将银行报盘文件（.xls）按单位名-年月-银行名称改名，同一单位不同银行合并为建行9列格式的 xlsx 文件，然后按合并文件排序顺序，通过 WPS COM 自动化批量打印对应的已签字工资表。
>
> **Deliverables**:
> - Python 主程序（GUI + 核心逻辑）
> - PowerShell 启动脚本（.ps1）
> - `uv` 虚拟环境配置
> - `.gitignore` 和 `README.md`
>
> **Estimated Effort**: Medium
> **Parallel Execution**: YES — Wave 1 (4 tasks), Wave 2 (2 tasks), Wave 3 (2 tasks + Final)
> **Critical Path**: Task 1 → Task 2 → Task 3 → Task 4 → Task 5 → Task 6 → Final QA

---

## Context

### Original Request
用户需要一个批量打印程序，处理银行报盘和工资表两部分文件：

1. **银行报盘**：按单位名-年月-银行名称改名，然后按单位名合并不同银行的 .xls → 一个 .xlsx，统一为建行9列格式
2. **工资表打印**：按合并后的银行报盘文件排序顺序，批量打印对应的 signed_ 工资表

### Interview Summary
**Key Discussions**:
- **技术栈**：Python + xlrd/openpyxl/pandas + GUI (tkinter)，用 uv 管理虚拟环境，.ps1 启动
- **列映射**：工商银行（3列：银行卡号/员工姓名/实发金额）→ 映射到建行9列格式，缺字段留空
- **跨行标识**：建设银行留空，工商银行填"1"，吉林银行原值
- **行名**：有原值用原值，没有按文件名填写（建设银行→"建设银行"，工商银行→"工商银行"，吉林银行→"吉林银行"）
- **打印**：WPS COM (`KET.Application`)，A4横向，每个文件一次打印，`signed_` 版本优先
- **排序**：Windows默认字符串升序
- **单位清单**：自动扫描目录匹配（97个单位 100% 匹配）

**Research Findings**:
- 样例数据在 `/home/ubuntu/excel_example/baiyun/`，银行报盘97个 .xls，3种银行格式
- 建设银行/吉林银行原始已是9列格式，工商银行仅3列
- 工资表有 `signed_` 和无前缀两个版本，3个 `.xls` 格式需 xlrd 处理
- 9个单位有多个银行（2-3家）

### Metis Review
**Identified Gaps** (addressed):
- WPS COM 打印机制：确认用 `KET.Application` COM 自动化
- 建设银行行名/跨行标识：行名填"建设银行"，跨行标识留空
- 混合括号文件名：`(吉林大学）深部探测...` — 用正则分割而非简单 split('-')
- 3个 `.xls` 工资表：xlrd 读取，而非 openpyxl
- 排序：Windows 默认字符串排序，括号保持原样

---

## Work Objectives

### Core Objective
一个带 GUI 的 Python 工具，将银行报盘文件按单位名-年月-银行名称改名，同一单位不同银行合并为统一9列格式的 xlsx 文件，然后按合并文件排序顺序，通过 WPS COM 自动化批量打印对应的已签字工资表。

### Concrete Deliverables
- `batchprint_gui.py` — 主程序（GUI + 核心逻辑）
- `run_batch_print.ps1` — PowerShell 启动脚本
- `pyproject.toml` — uv 项目配置
- `.gitignore`
- `README.md` — 使用说明

### Definition of Done
- [ ] 选择银行报盘目录 → 程序自动改名所有 .xls 文件到临时目录
- [ ] 选择工资表目录 → 程序匹配到所有 `signed_` 工资表
- [ ] 选择输出目录 → 程序合并生成 .xlsx 文件
- [ ] 点击"开始打印" → 97个工资表按顺序通过 WPS 打印

### Must Have
- 改名：原始 `202606-工商银行-吉林大学数学学院.xls` → `吉林大学数学学院-202606-工商银行.xls`
- 合并：同一单位多银行文件 → 一个 xlsx，建行9列 [序号	账户	户名	金额	跨行标识	行名	联行行号	摘要	备注]
- 合并文件名：`吉林大学数学学院-202606-工商银行2-建设银行1.xlsx`
- 银行按文件名中出现的字母顺序排列
- 打印顺序与合并文件按文件名升序排序一致
- 打印 `signed_` 版本的工资表（无 signed_ 则打印普通版本）
- GUI 界面让用户选择银行报盘/工资表/输出三个目录
- .ps1 脚本自动激活 uv venv 并启动程序

### Must NOT Have (Guardrails)
- ❌ 不修改原始银行报盘文件（改名操作只在临时目录/输出目录进行）
- ❌ 不生成 PDF 文件
- ❌ 不发送邮件或网络通信
- ❌ 不增加数据库存储
- ❌ 不处理多个月份的批量
- ❌ 不处理 xls/xlsx 以外的格式
- ❌ 不新增第4种银行类型的支持
- ❌ 人工智能/Slop：不要过度抽象（一个 py 文件即可，不要拆成多模块）；不要添加多余日志或打印信息；不要在 GUI 上加多余装饰

---

## Verification Strategy

> **ZERO HUMAN INTERVENTION** — ALL verification is agent-executed. No exceptions.

### Test Decision
- **Infrastructure exists**: Python (pytest available)
- **Automated tests**: Tests-after (核心逻辑单元测试)
- **Framework**: pytest + sample data at `/home/ubuntu/excel_example/baiyun/`
- **打印部分**: Windows机器上由用户手动验证（COM无法在Linux测试）

### QA Policy
对核心逻辑（改名、合并、列映射、匹配）使用样例数据进行验证。打印部分因依赖 Windows COM，仅做代码检查。
证据保存到 `.omo/evidence/task-{N}-{scenario-slug}.{ext}`。

---

## Execution Strategy

### Parallel Execution Waves

```
Wave 1 (Initialize - project scaffolding + core logic):
├── Task 1: 项目脚手架 + uv 配置 + .ps1 脚本 [quick]
├── Task 2: 核心逻辑模块 — 文件改名 + 列映射 + 合并 [unspecified-high]
├── Task 3: 工资表匹配模块 [unspecified-high]
└── Task 4: WPS 打印模块 [unspecified-high]

Wave 2 (GUI + Integration):
├── Task 5: GUI 界面 [visual-engineering]
└── Task 6: 测试 + 验证 [unspecified-low]

Wave FINAL:
├── Task F1: Plan compliance audit (oracle)
├── Task F2: Code quality + 样例数据验证 (unspecified-high)
├── Task F3: 打印模块代码审查 (unspecified-high)
└── Task F4: README (writing)
```

### Dependency Matrix

- **2**: None — 5
- **3**: None — 5
- **4**: None — 5
- **5**: 2, 3, 4 — 6
- **6**: 5 — F1-F4

### Agent Dispatch Summary

- **Wave 1**: 4 parallel — Task1(quick), Task2(unspecified-high), Task3(unspecified-high), Task4(unspecified-high)
- **Wave 2**: 1 task — Task5(visual-engineering), Task6(unspecified-low)
- **Final**: 4 parallel — F1-F4

---

## TODOs

- [x] 1. 项目脚手架 + uv 配置 + .ps1 启动脚本 + README

  **What to do**:
  - 创建 `pyproject.toml`，配置 uv 项目，依赖：xlrd, openpyxl, pandas, pywin32
  - 创建 `.gitignore`（忽略 `venv/`, `__pycache__/`, `.omo/`, `*.ps1` 除外）
  - 创建 `run_batch_print.ps1`：
    - 检测 `uv` 是否安装，否则提示安装
    - `uv sync` 同步依赖
    - `uv run python batchprint_gui.py`
    - 暂停等待用户按回车退出
  - 创建 `README.md`（中文，详细使用说明）

  **Must NOT do**:
  - 不要用 `pip`，必须用 `uv`
  - 不要硬编码路径

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 标准化的脚手桑搭建工作
  - **Skills**: 无需额外技能

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3, 4)
  - **Blocks**: None
  - **Blocked By**: None (can start immediately)

  **Acceptance Criteria**:
  - [ ] `pyproject.toml` 存在且包含 xlrd, openpyxl, pandas, pywin32 依赖
  - [ ] `pyproject.toml` 包含 `[project.scripts]` 入口点
  - [ ] `.gitignore` 存在且匹配预期内容
  - [ ] `run_batch_print.ps1` 存在且语法正确
  - [ ] `README.md` 存在且包含使用说明

  **QA Scenarios**:
  ```
  Scenario: Check project files exist and are valid
    Tool: Bash
    Preconditions: Working directory at /home/ubuntu/coding/seqprint
    Steps:
      1. Check `ls pyproject.toml .gitignore run_batch_print.ps1 README.md` all exist
      2. Validate `pyproject.toml` is valid TOML: `python3 -c "import tomllib; tomllib.load(open('pyproject.toml','rb'))"`
      3. Validate `run_batch_print.ps1` has correct uv commands
    Expected Result: All 4 files exist, TOML is valid, ps1 has uv run pattern
    Evidence: .omo/evidence/task-1-files-check.txt
  ```

  **Commit**: YES (groups with 2-4)
  - Message: `feat: project scaffold with uv config and startup script`
  - Files: `pyproject.toml`, `.gitignore`, `run_batch_print.ps1`, `README.md`

---

- [x] 2. 核心逻辑模块 — 文件改名 + 列映射 + 合并

  **What to do**:
  在 `batchprint_gui.py` 中实现以下函数（在同一个文件中，不要分文件）：

  **函数一: rename_bank_files(bank_dir, output_dir)**
  - 扫描 `bank_dir` 下所有 `.xls` 文件
  - 原始文件名格式：`YYYYMM-银行名称-单位名.xls`
  - 改名规则：`单位名-YYYYMM-银行名称.xls`
  - **注意不能用简单 split('-')**：因为 `(吉林大学）深部探测与成像全国重点实验室` 这类单位名可能含 `-`
  - 文件名只包含一个银行名，所以分割策略：第一个 `-` 前是年月，第二个 `-` 前是银行名，之后全是单位名
  - 复制到 `output_dir`（不修改源文件）
  - 返回重命名后文件列表 `[(new_name, old_name, bank_name, unit_name), ...]`

  **.二: merge_bank_files(renamed_list, bank_dir, output_dir)**
  - 按单位名分组 → 每个单位一个合并文件
  - 建设银行和吉林银行：原始文件是9列格式 `[序号, 账户, 户名, 金额, 跨行标识, 行名, 联行行号, 摘要, 备注]`，直接复制
  - 工商银行：3列 `[银行卡号, 员工姓名, 实发金额]` → 映射到9列
  - 列映射规则：
    | 建行列 | 工行源 | 建行源 | 吉林源 |
    |--------|--------|--------|--------|
    | 序号 | 1,2,3... | 复制 | 复制 |
    | 账户 | 银行卡号 | 复制 | 复制 |
    | 户名 | 员工姓名 | 复制 | 复制 |
    | 金额 | 实发金额 | 复制 | 复制 |
    | 跨行标识 | "1" | 留空 | 复制 |
    | 行名 | "工商银行" | "建设银行" | 有值用值，无则"吉林银行" |
    | 联行行号 | "" | 复制 | 复制 |
    | 摘要 | "" | 复制 | 复制 |
    | 备注 | "" | 复制 | 复制 |
  - 银行在文件名中的顺序：按字母排序（建设银行 < 工商银行 < 吉林银行）
  - 合并文件名：`{单位名}-{年月}-{银行1}{行数}-{银行2}{行数}.xlsx`
  - 银行内不排序，保持原有行顺序

  **.三: 辅助函数 detect_bank_type(filename)** — 从文件名中的银行名判断
  **.四: 辅助函数 split_filename(filename)** — 正确解析 `年月-银行-单位名` 三部分

  **Must NOT do**:
  - 不要创建多个 .py 文件，所有代码在 `batchprint_gui.py` 中
  - 不要使用 `pathlib`，用 `os.path` 即可
  - 不要添加多余的类型注解（除非必要）
  - 不要过度错误处理（顶级 try/except 即可）

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: 无需额外技能

  **Parallelization**:
  - **Can Run In Parallel**: YES (with Task 1, 3, 4)
  - **Parallel Group**: Wave 1
  - **Blocks**: Task 5 (GUI depends on core logic)
  - **Blocked By**: None

  **Acceptance Criteria**:
  - [ ] `rename_bank_files()` 能正确解析 `202606-建设银行-吉林大学数学学院.xls` → 三部分分离
  - [ ] `rename_bank_files()` 能正确解析带括号的文件（如含 `-` 的单位名）
  - [ ] 工商银行 3列 → 正确映射为建行9列
  - [ ] 建设银行/吉林银行 9列 → 正确复制
  - [ ] 合并后每家银行的行数正确
  - [ ] 多银行单位的合并文件名正确：`吉林大学数学学院-202606-工商银行2-建设银行1.xlsx`
  - [ ] 单银行单位的合并文件名正确：`吉林大学物理学院-202606-建设银行1.xlsx`

  **QA Scenarios**:
  ```
  Scenario: Test rename with sample bank files
    Tool: Bash
    Preconditions: pytest available, sample data at /home/ubuntu/excel_example/baiyun/银行报盘/
    Steps:
      1. Create test script that imports parse_filename() from batchprint_gui.py
      2. Test parse: "202606-建设银行-吉林大学数学学院.xls" → ("202606", "建设银行", "吉林大学数学学院")
      3. Test parse: "202606-建设银行-(吉林大学）深部探测与成像全国重点实验室.xls" → correct 3 parts
      4. Create temporary test directory, run rename_bank_files, check output filenames
    Expected Result: All filenames correctly parsed and renamed
    Evidence: .omo/evidence/task-2-rename-test.txt

  Scenario: Test ICBC to CCB column mapping
    Tool: Bash
    Preconditions: ICBC sample .xls files available
    Steps:
      1. Read an ICBC file with xlrd, confirm 3 original columns
      2. After merge: confirm 9 output columns
      3. Check row count matches
      4. Check non-ICBC columns are empty or bank name filled
    Expected Result: 3→9 column mapping correct, bank name "工商银行" filled in 行名
    Evidence: .omo/evidence/task-2-icbc-mapping.txt
  ```

  **Commit**: YES (groups with 1, 3, 4)
  - Message: `feat: bank file rename, merge and column mapping logic`
  - Files: `batchprint_gui.py` (partial)

---

- [x] 3. 工资表匹配模块

  **What to do**:
  在 `batchprint_gui.py` 中实现以下函数：

  **.一: match_payroll_files(merged_files_list, payroll_dir)**
  - 输入：合并后的银行报盘文件列表（已排序）和工资表目录
  - 对每个合并文件，提取单位名（`X-202606-银行A-银行B.xlsx` 中的 `X`）
  - 在 `payroll_dir` 下查找对应工资表：
    1. 优先找 `signed_{单位名}...xlsx`（匹配包含单位名的 signed_ 文件）
    2. 若无 signed_，则找普通 `{单位名}...xlsx`
    3. 若还是没有，则找 `.xls` 版本（用 xlrd）
  - 单位名匹配用 `in` 包含关系（单位名是工资表文件名的子串）
  - 排除 `汇总表及验证明细.xlsx` 文件
  - 返回：`[(merged_filename, payroll_filepath, unit_name), ...]` — 与合并文件列表**一一对应**

  **R5: 模糊匹配建议用 "单位名 in 工资表文件名" 方式**，因为样例如：
  - 合并文件名：`吉林大学数学学院-202606-工商银行2-建设银行1.xlsx`
  - 工资表：`signed_吉林大学数学学院2026年06月工资表.xlsx`
  - 单位名 = `吉林大学数学学院` → 在 `signed_吉林大学数学学院2026年06月工资表.xlsx` 中找到

  **Must NOT do**:
  - 不要用汇总表做匹配（用户确认了自动扫描）
  - 不要改工资表文件

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Reason**: 文件匹配逻辑需要处理多种命名模式
  - **Skills**: 无需额外技能

  **Parallelization**:
  - **Can Run In Parallel**: YES (with Tasks 1, 2, 4)
  - **Parallel Group**: Wave 1
  - **Blocks**: Task 5 (GUI needs match results)
  - **Blocked By**: None

  **Acceptance Criteria**:
  - [ ] 对所有 97 个单位，每个都能匹配到对应的工资表文件
  - [ ] `signed_` 版本优先
  - [ ] 返回列表长度与合并文件列表一致
  - [ ] `.xls` 工资表正确处理
  - [ ] `汇总表` 文件被排除

  **QA Scenarios**:
  ```
  Scenario: Test payroll matching with sample directory
    Tool: Bash
    Preconditions: Merged files list created from Task 2, payroll_dir = /home/ubuntu/excel_example/baiyun/.../当天工资/... 
    Steps:
      1. Create test script with match_payroll_files()
      2. Pass 97 merged filenames, check all 97 matches found
      3. Verify all matched files are signed_ when available
      4. Check only 汇总表 excluded
    Expected Result: 97/97 matches, signed_ preferred, 汇总表 excluded
    Evidence: .omo/evidence/task-3-matching.txt
  ```

  **Commit**: YES (groups with 1, 2, 4)
  - Message: `feat: payroll file matching module`
  - Files: `batchprint_gui.py` (partial)

---

- [x] 4. WPS 打印模块

  **What to do**:
  在 `batchprint_gui.py` 中实现以下函数：

  **.6: print_file(filepath)**
  - 使用 `win32com.client.Dispatch("KET.Application")` 打开 WPS
  - 打开 `filepath` 的 Excel 文件
  - 设置页面：A4横向
  - 打印当前活动工作表（`ActiveSheet.PrintOut()`）
  - 关闭文件（不保存），不退出 WPS 应用（保留以备多次打印）
  - 每个文件结束后等待 2-3 秒（避免打印队列堵塞）
  - 返回 True/False 表示成功/失败
  - 如果打印失败，重试最多 3 次，每次间隔 5 秒

  **.7: batch_print(matched_pairs)**
  - 依次遍历 `matched_pairs` 列表（已排序）
  - 对每个文件调用 `print_file()`
  - 在 GUI 中更新进度（打印到第 N/97 个）
  - 打印失败时记录到结果列表，不影响后续打印
  - 所有打印完成后，汇总显示：成功数/失败数/失败文件名列表

  **.8: check_wps_available()**
  - 尝试 `Dispatch("KET.Application")`
  - 失败则返回 False，GUI 显示"WPS Office 未安装或 COM 组件不可用"

  **Must NOT do**:
  - 不要试图关闭已打开的 WPS 窗口
  - 不要修改原始文件

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Reason**: WPS COM 自动化需要正确的 COM 调用模式
  - **Skills**: 无需额外技能

  **Parallelization**:
  - **Can Run In Parallel**: YES (with Tasks 1, 2, 3)
  - **Parallel Group**: Wave 1
  - **Blocks**: Task 5 (GUI needs print functions)
  - **Blocked By**: None

  **Acceptance Criteria**:
  - [ ] 代码中包含 `check_wps_available()` 函数
  - [ ] `print_file()` 使用 `KET.Application` COM 接口
  - [ ] 打印设置包含 A4 横向
  - [ ] 有重试机制（最多3次）
  - [ ] 有进度反馈和结果汇总
  - [ ] 代码风格与核心逻辑一致

  **QA Scenarios**:
  ```
  Scenario: Code review of WPS print module (Linux can't run COM)
    Tool: Bash
    Preconditions: batchprint_gui.py contains print module
    Steps:
      1. Check win32com.client.Dispatch("KET.Application") is called
      2. Check PageSetup.Orientation = 2 (xlLandscape)
      3. Check PrintOut is called on the active sheet
      4. Check retry loop (3 attempts)
    Expected Result: Correct COM API calls for WPS printing
    Evidence: .omo/evidence/task-4-code-review.txt
  ```

  **Commit**: YES (groups with 1, 2, 3)
  - Message: `feat: WPS COM print module`
  - Files: `batchprint_gui.py` (partial)

---

- [x] 5. GUI 界面

  **What to do**:
  在 `batchprint_gui.py` 中实现 GUI 主程序（使用 tkinter），调用 Tasks 2, 3, 4 的函数：

  **GUI 布局**：
  ```
  标题：批量银行报盘合并与工资表打印工具

  [选择银行报盘目录] [选择工资表目录] [选择输出目录]
  ─────────────────────────────────────────────
  银行报盘目录：/path/to/银行报盘/
  工资表目录：/path/to/工资表/
  输出目录：/path/to/输出/

  ┌─ 日志信息 ──────────────────────────────┐
  │ [步骤1] 改名完成：97个文件                │
  │ [步骤2] 合并完成：97个合并文件            │
  │ [步骤3] 匹配完成：97/97 匹配成功          │
  │ [步骤4] 打印进度：23/97                   │
  └──────────────────────────────────────────┘

  [执行全部并打印] [仅合并不打印]
  ```

  **功能**：
  - 3个目录选择按钮 → 弹出 `filedialog.askdirectory()`
  - "执行全部并打印"按钮：
    1. 调用 `rename_bank_files()` + 显示进度
    2. 调用 `merge_bank_files()` + 显示进度
    3. 调用 `match_payroll_files()` + 显示匹配结果
    4. 检查 WPS 可用
    5. 展示合并文件列表（滚动文本框）
    6. 用户确认后开始打印 → 调用 `batch_print()`
  - "仅合并不打印"按钮：执行步骤 1-3，输出合并文件，不打印
  - 日志区域用 `ScrolledText`，每一步追加日志信息
  - 执行时按钮禁用，完成后启用

  **Must NOT do**:
  - 不要用 webview/html/css 等（只用 tkinter）
  - 不要太花哨：标准 tkinter 控件即可
  - 不要设置窗口图标

  **Recommended Agent Profile**:
  - **Category**: `visual-engineering`
  - **Reason**: 需要 tkinter GUI 设计
  - **Skills**: 无需额外技能

  **Parallelization**:
  - **Can Run In Parallel**: NO (depends on 2, 3, 4)
  - **Parallel Group**: Wave 2 (sequential with Task 6)
  - **Blocks**: Task 6 (testing)
  - **Blocked By**: Tasks 2, 3, 4

  **Acceptance Criteria**:
  - [ ] 三个目录选择按钮都工作
  - [ ] "执行全部并打印"触发完整流程
  - [ ] "仅合并不打印"触发步骤 1-3
  - [ ] 日志区域显示实时进度
  - [ ] 执行中按钮禁用
  - [ ] 窗口标题为"批量打印程序"

  **QA Scenarios**:
  ```
  Scenario: Test GUI launches
    Tool: Bash
    Preconditions: batchprint_gui.py exists with all tasks
    Steps:
      1. Python syntax check: `python3 -c "import ast; ast.parse(open('batchprint_gui.py').read())"`
      2. Check tkinter import exists
      3. Check askdirectory() is used for all 3 directory inputs
      4. Check mainloop() or main() entry point exists
    Expected Result: Valid Python syntax, tkinter used, directory choosers present
    Evidence: .omo/evidence/task-5-gui-check.txt
  ```

  **Commit**: YES (groups with 6)
  - Message: `feat: GUI interface for batch print tool`
  - Files: `batchprint_gui.py` (complete)

---

- [x] 6. 测试 + 验证

  **What to do**:
  使用样例数据验证核心逻辑的正确性：

  **验证 1: 改名测试**
  - 复制银行报盘样例目录到临时目录
  - 调用 `rename_bank_files()`，验证输出文件名格式正确
  - 验证文件名解析正确（含括号的特殊文件）

  **验证 2: 合并列映射测试**
  - 对工商银行文件调用 `merge_bank_files()`
  - 验证 3 列 → 9 列映射正确
  - 验证行名 / 跨行标识填充正确

  **验证 3: 匹配测试**
  - 调用 `match_payroll_files()`
  - 验证 97 个单位都能匹配到工资表文件
  - 验证优先选择 `signed_` 版本

  **验证 4: 排序测试**
  - 验证合并文件按文件名升序排列
  - 验证打印顺序与此一致

  **不要求创建独立的 `test_` 文件**，可以用临时测试脚本运行验证。

  **Must NOT do**:
  - 不要写 `pytest` 测试套件（仅验证性测试）
  - 不要修改原始样例文件

  **Recommended Agent Profile**:
  - **Category**: `unspecified-low`
  - **Reason**: 验证性测试，逻辑简单

  **Parallelization**:
  - **Can Run In Parallel**: NO (depends on Task 5 complete)
  - **Parallel Group**: Wave 2 (after Task 5)
  - **Blocks**: None
  - **Blocked By**: Task 5

  **Acceptance Criteria**:
  - [ ] 改名测试通过：97个文件全部改名正确
  - [ ] 列映射测试通过：工商银行3→9列映射正确
  - [ ] 匹配测试通过：97/97 匹配成功
  - [ ] 排序测试通过：合并文件按 Windows 默认升序

  **QA Scenarios**:
  ```
  Scenario: Full integration test with sample data
    Tool: Bash
    Preconditions: Sample data in /home/ubuntu/excel_example/baiyun/
    Steps:
      1. Create temp test directories
      2. Run rename test with 97 files
      3. Run merge test, verify output filenames
      4. Run match test with payroll files
      5. Validate 9-column format for merged output
    Expected Result: All 97 files processed correctly
    Evidence: .omo/evidence/task-6-integration.txt
  ```

  **Commit**: NO (groups with 5)
  - Message: (same as task 5 commit)
  - Files: `batchprint_gui.py` (final)

---

## Final Verification Wave (after ALL implementation tasks)

- [x] F1. **Plan Compliance Audit** — `oracle`
  Read `batchprint_gui.py` end to end. Verify: Must Have all present (rename, merge to 9 columns, WPS print, GUI, signed_ matching). Must NOT Have all absent (no extra modules, no original file modification). Check evidence files exist.
  Output: `Must Have [N/N] | Must NOT Have [N/N] | Tasks [N/N] | VERDICT: APPROVE/REJECT`

- [x] F2. **Code Quality + 样例验证** — `unspecified-high`
  Run Python syntax check on `batchprint_gui.py`. Review for: type suppressions, empty catches, debug prints, over-abstraction. Use sample data to run rename+merge+match logic, validate outputs.
  Output: `Syntax [PASS/FAIL] | Core logic [N scenarios pass/N fail] | Code quality [clean/issues] | VERDICT`

- [x] F3. **WPS 打印模块审查** — `unspecified-high`
  Code review the print module: verify COM API calls are syntactically correct (KET.Application, PrintOut, PageSetup). 验证重试逻辑、错误处理、进度回调。
  Output: `COM API [correct/incorrect] | Retry [N] | Error handling [PASS/FAIL] | VERDICT`

- [x] F4. **README + 文档** — `writing`
  检查中文 README 是否包含：功能说明、运行环境要求（Windows + WPS）、安装步骤（uv）、使用截图指导。
  Output: `Sections [N/N required] | Clarity [PASS/FAIL] | VERDICT`

---

## Commit Strategy

- **Commit 1** (Tasks 1-4): `feat: project scaffold, core logic, matching and print modules`
  - Files: `pyproject.toml`, `.gitignore`, `run_batch_print.ps1`, `README.md`, `batchprint_gui.py`
  
- **Commit 2** (Tasks 5-6): `feat: GUI interface and integration testing`
  - Files: `batchprint_gui.py` (final)

---

## Success Criteria

### Verification Commands
```bash
# Syntax check
python3 -c "import ast; ast.parse(open('batchprint_gui.py').read())" && echo "Syntax OK"

# Sample data validation
python3 -c "
import sys; sys.path.insert(0, '.')
# Import and run core validation
from batchprint_gui import rename_bank_files, merge_bank_files, match_payroll_files
# ... run all validation steps ...
"
```

### Final Checklist
- [ ] 项目文件全部存在：`pyproject.toml`, `.gitignore`, `run_batch_print.ps1`, `README.md`, `batchprint_gui.py`
- [ ] Python 语法正确
- [ ] 核心逻辑通过样例数据验证（改名/合并/匹配）
- [ ] 97个单位全部成功匹配
- [ ] 列映射正确（工商银行3→9列）
- [ ] GUI 界面包含所有必需元素
- [ ] WPS 打印模块使用正确的 COM API
- [ ] .ps1 启动脚本包含 uv 环境激活
- [ ] README 包含 Windows 运行说明
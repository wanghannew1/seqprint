# 批量银行报盘合并与工资表打印工具

## 功能说明

1. **改名银行报盘文件** — 自动识别并重命名银行报盘文件
2. **合并为建设银行 9 列格式** — 将多个银行报盘合并为统一格式
3. **批量打印工资表** — 自动打开 WPS Office 批量打印工资表

## 运行环境要求

- **操作系统**: Windows 10/11
- **办公软件**: WPS Office（用于打印功能）
- **Python**: 3.10 或更高版本
- **包管理器**: uv（Python 包管理器）

## 安装步骤

### 方式一：使用 uv（推荐）

1. 安装 uv（如未安装）：
   ```powershell
   powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
   ```

2. 双击 `run_batch_print.ps1` 启动，脚本会自动安装依赖并运行。
   （pyproject.toml 已配置清华镜像源，下载速度更快）

### 方式二：使用 uv pip（已有 uv 环境时）

```powershell
uv pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
python batchprint_gui.py
```

### 已验证环境
- Python 3.13 + uv（清华镜像）：✅ 通过
- 依赖：xlrd, openpyxl, pandas, pywin32
- pywin32 仅在 Windows 下自动安装

## 使用说明

1. **选择银行报盘目录** — 包含各银行报盘 Excel 文件的文件夹
2. **选择工资表目录** — 包含工资表模板的文件夹
3. **选择输出目录** — 合并结果和打印文件的保存位置
4. **点击执行** — 程序自动完成改名、合并和打印

## 注意事项

- 首次运行 `uv sync` 会自动创建虚拟环境并安装依赖
- 如遇网络慢，pyproject.toml 已默认配置清华镜像
- 程序不会修改原始文件，所有操作均在副本上进行
- 仅支持以下银行格式：
  - 中国建设银行
  - 中国工商银行
  - 吉林银行
- 打印功能依赖 WPS Office，请确保已安装
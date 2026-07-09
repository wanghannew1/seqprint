# 批量打印程序启动脚本
Write-Host "=== 批量银行报盘合并与工资表打印工具 ===" -ForegroundColor Cyan

# 检查 uv 是否安装
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Host "错误: 未检测到 uv，请先安装: https://docs.astral.sh/uv/" -ForegroundColor Red
    Write-Host "安装命令: powershell -c `"irm https://astral.sh/uv/install.ps1 | iex`"" -ForegroundColor Yellow
    Read-Host "按回车退出"
    exit 1
}

# 同步依赖
Write-Host "正在同步依赖..." -ForegroundColor Yellow
uv sync
if ($LASTEXITCODE -ne 0) {
    Write-Host "依赖安装失败" -ForegroundColor Red
    Read-Host "按回车退出"
    exit 1
}

# 启动程序
Write-Host "正在启动程序..." -ForegroundColor Green
uv run python batchprint_gui.py

# 暂停
Read-Host "按回车退出"
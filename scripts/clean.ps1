<#
.SYNOPSIS
    清理「批量发送邮件工具」项目里可再生的目录与文件。

.DESCRIPTION
    只删除**可以重新生成**的东西，源码、测试、脚本、assets、.github、verchanglog、
    dist 与 .venv 一律不动。

    默认是「预演」模式：只列出将要删除的内容，不实际删除。
    确认无误后加 -Execute 真正执行。

    安全设计：
      - 通过 scripts\.project-root 标记文件确认项目根目录，标记不存在就拒绝执行；
      - 只处理白名单内的路径，不做通配符递归删除；
      - 每删一项都打印出来，便于核对。

.PARAMETER Execute
    真正执行删除。不加此参数时仅预演。

.PARAMETER KeepWork
    保留 work\ 目录（本地验证脚本与临时产物，便于继续排查）。

.EXAMPLE
    .\scripts\clean.ps1
    预演：列出将被删除的内容。

.EXAMPLE
    .\scripts\clean.ps1 -Execute
    真正执行清理。

.EXAMPLE
    .\scripts\clean.ps1 -Execute -KeepWork
    清理但保留 work\ 目录。
#>
[CmdletBinding()]
param(
    [switch]$Execute,
    [switch]$KeepWork
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$marker = Join-Path $PSScriptRoot ".project-root"

if (-not (Test-Path -LiteralPath $marker)) {
    throw "未找到项目标记文件：$marker。为避免误删，拒绝执行。"
}

# 白名单：每一项都是「删掉后能重新生成」的东西
$targets = [System.Collections.Generic.List[object]]::new()

function Add-Target {
    param([string]$RelativePath, [string]$Kind, [string]$Reason)
    $full = Join-Path $projectRoot $RelativePath
    if (Test-Path -LiteralPath $full) {
        $targets.Add([PSCustomObject]@{
            Path   = $full
            Kind   = $Kind
            Reason = $Reason
        })
    }
}

# ① 打包中间产物（PyInstaller 每次构建都会重建）
Add-Target "build" "目录" "PyInstaller 中间产物，重新打包即重建"
Get-ChildItem -LiteralPath $projectRoot -File -Filter "*.spec" -ErrorAction SilentlyContinue |
    ForEach-Object {
        $targets.Add([PSCustomObject]@{
            Path   = $_.FullName
            Kind   = "文件"
            Reason = "打包脚本自动生成的 PyInstaller 配置"
        })
    }

# ② Python 与测试缓存（下次运行自动重建）
Add-Target ".pytest_cache" "目录" "pytest 缓存"
Add-Target ".coverage" "文件" "覆盖率数据文件"
Add-Target "src\send_email_studio.egg-info" "目录" "可编辑安装元数据，重装即重建"

# src 与 tests 下的缓存逐个包处理（不用通配符递归，避免误伤）
foreach ($relative in @("src\send_email_studio", "tests")) {
    $packageRoot = Join-Path $projectRoot $relative
    if (Test-Path -LiteralPath $packageRoot) {
        Get-ChildItem -LiteralPath $packageRoot -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
            ForEach-Object {
                $targets.Add([PSCustomObject]@{
                    Path   = $_.FullName
                    Kind   = "目录"
                    Reason = "Python 字节码缓存"
                })
            }
    }
}

# ③ 运行与调试残留
Add-Target "smoke.xlsx" "文件" "试发邮件用的临时表格"
Add-Target "work\dsh-session-current.md" "文件" "会话导出中间文件"
Get-ChildItem -LiteralPath $projectRoot -File -Filter "*.partial.*" -ErrorAction SilentlyContinue |
    ForEach-Object {
        $targets.Add([PSCustomObject]@{
            Path   = $_.FullName
            Kind   = "文件"
            Reason = "中断的导出产物"
        })
    }
# 排查问题时把测试输出重定向到 work\ 会产生这些日志
Get-ChildItem -LiteralPath (Join-Path $projectRoot "work") -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "*.log" -or $_.Name -like "*.err" } |
    ForEach-Object {
        $targets.Add([PSCustomObject]@{
            Path   = $_.FullName
            Kind   = "文件"
            Reason = "排查用的临时日志"
        })
    }

# ④ work\ 目录本身（本地验证脚本与产物）
if (-not $KeepWork) {
    Add-Target "work" "目录" "本地验证脚本与临时产物，加 -KeepWork 可保留"
}

if ($targets.Count -eq 0) {
    Write-Host "没有需要清理的内容，项目目录已经很干净。" -ForegroundColor Green
    exit 0
}

# 统计大小
function Get-TargetSize {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        return (Get-Item -LiteralPath $Path).Length
    }
    return (Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum
}

$mode = if ($Execute) { "执行清理" } else { "预演（不会删除任何东西）" }
Write-Host ""
Write-Host "项目根目录：$projectRoot"
Write-Host "运行模式　：$mode" -ForegroundColor $(if ($Execute) { "Yellow" } else { "Cyan" })
Write-Host ""
Write-Host ("{0,-9} {1,10}  {2}" -f "类型", "大小", "路径 / 说明")
Write-Host ("-" * 96)

$totalBytes = 0
foreach ($target in $targets) {
    $size = Get-TargetSize -Path $target.Path
    $totalBytes += $size
    $relative = $target.Path.Substring($projectRoot.Length).TrimStart("\")
    $sizeText = if ($size -ge 1MB) { "{0:N2} MB" -f ($size / 1MB) } else { "{0:N1} KB" -f ($size / 1KB) }
    Write-Host ("{0,-9} {1,10}  {2}" -f $target.Kind, $sizeText, $relative)
    Write-Host ("{0,-9} {1,10}    └─ {2}" -f "", "", $target.Reason) -ForegroundColor DarkGray
}

Write-Host ("-" * 96)
$totalText = if ($totalBytes -ge 1MB) { "{0:N2} MB" -f ($totalBytes / 1MB) } else { "{0:N1} KB" -f ($totalBytes / 1KB) }
Write-Host ("共 {0} 项，合计 {1}" -f $targets.Count, $totalText)

if (-not $Execute) {
    Write-Host ""
    Write-Host "以上仅为预演。确认无误后执行：" -ForegroundColor Cyan
    Write-Host "    .\scripts\clean.ps1 -Execute" -ForegroundColor White
    exit 0
}

Write-Host ""
Write-Host "开始清理…" -ForegroundColor Yellow
$removed = 0
$failed = 0
foreach ($target in $targets) {
    try {
        Remove-Item -LiteralPath $target.Path -Recurse -Force -ErrorAction Stop
        $removed++
    }
    catch {
        $failed++
        Write-Warning ("删除失败：{0} —— {1}" -f $target.Path, $_.Exception.Message)
    }
}

Write-Host ""
Write-Host ("清理完成：成功 {0} 项，失败 {1} 项，释放 {2}" -f $removed, $failed, $totalText) -ForegroundColor Green
if ($failed -gt 0) {
    Write-Host "有项目删除失败，通常是文件被占用（先关掉正在运行的工具或 Excel）。" -ForegroundColor Yellow
    exit 1
}

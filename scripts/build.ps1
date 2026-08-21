param(
    [switch]$OneFile,
    [string]$Name = "批量发送邮件工具"
)

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$icon = Join-Path $projectRoot "assets\icons\app.ico"

if (-not (Test-Path -LiteralPath $python)) {
    throw "未找到项目虚拟环境，请先创建 .venv 并安装依赖。"
}

$arguments = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--name", $Name,
    "--windowed",
    "--paths", "src",
    "scripts\pyinstaller_entry.py"
)

if (Test-Path -LiteralPath $icon) {
    $arguments += @("--icon", $icon, "--add-data", "$icon;assets\icons")
}
else {
    Write-Warning "未提供自定义图标，将使用默认应用图标。"
}

if ($OneFile) {
    $arguments += "--onefile"
}

Push-Location $projectRoot
try {
    & $python @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller 打包失败，退出码：$LASTEXITCODE"
    }
}
finally {
    Pop-Location
}


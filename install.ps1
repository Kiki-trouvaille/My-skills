# 把本仓库的技能同步到 WorkBuddy 用户级技能目录
# 用法：
#   .\install.ps1                         同步全部技能
#   .\install.ps1 -Skill group-meeting-ppt  只同步一个
#   .\install.ps1 -List                   只列出可用技能
# 注意：使用 robocopy /MIR 镜像同步，会删除目标目录中源目录没有的文件。

param(
    [string[]]$Skill = @(),
    [string]$Target = $(Join-Path $env:USERPROFILE ".workbuddy\skills"),
    [switch]$List
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$encoding = [Console]::OutputEncoding
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$available = Get-ChildItem -Path $root -Directory |
    Where-Object { Test-Path (Join-Path $_.FullName "SKILL.md") } |
    Select-Object -ExpandProperty Name

if (-not $available) {
    Write-Host "没有找到含 SKILL.md 的技能目录。" -ForegroundColor Yellow
    exit 1
}

if ($List) {
    Write-Host "可用技能："
    $available | ForEach-Object { Write-Host "  - $_" }
    exit 0
}

if (-not $Skill) { $Skill = $available }

$failed = @()
foreach ($name in $Skill) {
    $src = Join-Path $root $name
    if (-not (Test-Path (Join-Path $src "SKILL.md"))) {
        Write-Host "跳过 ${name}：目录不存在或缺少 SKILL.md" -ForegroundColor Yellow
        $failed += $name
        continue
    }

    $dst = Join-Path $Target $name
    New-Item -ItemType Directory -Force -Path $dst | Out-Null

    # /MIR 镜像；排除版本库与缓存，避免把 .git 和 __pycache__ 带进技能目录
    $null = robocopy $src $dst /MIR /XD ".git" "__pycache__" ".venv" /NFL /NDL /NJH /NJS /NP
    if ($LASTEXITCODE -ge 8) {
        Write-Host "失败 ${name}：robocopy 返回 $LASTEXITCODE" -ForegroundColor Red
        $failed += $name
    }
    else {
        Write-Host "已安装 ${name}  ->  ${dst}" -ForegroundColor Green
    }
}

[Console]::OutputEncoding = $encoding
if ($failed) {
    Write-Host ""
    Write-Host "未完成：$($failed -join ', ')" -ForegroundColor Red
    exit 1
}
Write-Host ""
Write-Host "完成。重启会话或新开对话后技能生效。"

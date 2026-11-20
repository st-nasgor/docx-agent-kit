#Requires -Version 5.1
<#
.SYNOPSIS
Install the docx-agent-kit skill into an agent's skills directory.

.DESCRIPTION
The repository can live anywhere. This script copies the skill into the
skills directory OpenCode or Claude Code already reads, then installs the
Python package from that installed copy.

.PARAMETER Agent
auto (default), opencode, or claude. Auto prefers an existing OpenCode
config, then Claude Code.

.PARAMETER SkillsDir
Explicit skills directory, for example .opencode/skills for a project.
Overrides -Agent.

.PARAMETER Link
Link the repository instead of copying it. Useful while developing the kit.
#>
param(
    [ValidateSet("auto", "opencode", "claude")]
    [string]$Agent = "auto",
    [string]$SkillsDir = "",
    [switch]$Link,
    # Also install Anthropic's document skills (docx, pdf, pptx, xlsx) with
    # the Windows patches and Python dependencies they need.
    [switch]$WithOfficeSkills
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "docx-agent-kit installer"
Write-Host "  source: $root"

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    throw "Python 3.10+ is required but 'python' was not found on PATH."
}
$version = & python -c "import sys; print('.'.join(map(str, sys.version_info[:2])))"
if ([version]$version -lt [version]"3.10") {
    throw "Python 3.10+ is required (found $version)."
}
Write-Host "  python: $version"

function Get-AgentSkillsDir {
    param([string]$Requested)

    if ($Requested -and $Requested -ne "auto") {
        switch ($Requested) {
            "opencode" { return (Join-Path $(if ($env:XDG_CONFIG_HOME) { $env:XDG_CONFIG_HOME } else { Join-Path $env:USERPROFILE ".config" }) "opencode\skills") }
            "claude"   { return (Join-Path $env:USERPROFILE ".claude\skills") }
        }
    }

    $opencodeConfig = Join-Path $(if ($env:XDG_CONFIG_HOME) { $env:XDG_CONFIG_HOME } else { Join-Path $env:USERPROFILE ".config" }) "opencode"
    $claudeConfig = Join-Path $env:USERPROFILE ".claude"
    if (Test-Path $opencodeConfig) { return (Join-Path $opencodeConfig "skills") }
    if (Test-Path $claudeConfig) { return (Join-Path $claudeConfig "skills") }
    return (Join-Path $opencodeConfig "skills")
}

if (-not $SkillsDir) { $SkillsDir = Get-AgentSkillsDir $Agent }
$SkillsDir = [System.IO.Path]::GetFullPath($SkillsDir)
$dest = Join-Path $SkillsDir "docx-agent-kit"

Write-Host "  skills: $SkillsDir"
New-Item -ItemType Directory -Force -Path $SkillsDir | Out-Null

if (Test-Path $dest) {
    $item = Get-Item $dest -Force
    if ($item.LinkType) {
        [System.IO.Directory]::Delete($dest)
    } else {
        Write-Host "  replacing existing $dest"
        Remove-Item $dest -Recurse -Force
    }
}

if ($Link) {
    try {
        New-Item -ItemType Junction -Path $dest -Target $root | Out-Null
        $installFrom = $root
        Write-Host "  linked: $dest -> $root"
    } catch {
        Write-Host "  link failed; copying instead"
        $Link = $false
    }
}

if (-not $Link) {
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    foreach ($name in @("SKILL.md", "README.md", "pyproject.toml", "docx_agent_kit", "tests")) {
        Copy-Item (Join-Path $root $name) -Destination $dest -Recurse -Force
    }
    Get-ChildItem $dest -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem $dest -Directory -Filter "*.egg-info" -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    $installFrom = $dest
    Write-Host "  installed: $dest"
}

Write-Host "  installing Python package..."
& python -m pip install -e $installFrom
if ($LASTEXITCODE -ne 0) { throw "pip install failed." }

& python -c "import docx_agent_kit; print('  import ok: docx_agent_kit', docx_agent_kit.__version__)"

if ($WithOfficeSkills) {
    Write-Host ""
    Write-Host "  --- office skills (docx, pdf, pptx, xlsx from anthropics/skills) ---"

    $clone = Join-Path $env:TEMP "anthropic-skills"
    if (Test-Path $clone) {
        Write-Host "  updating anthropics/skills clone..."
        git -C $clone pull --ff-only 2>$null | Out-Null
    } else {
        Write-Host "  cloning anthropics/skills..."
        git clone --depth 1 https://github.com/anthropics/skills.git $clone | Out-Null
    }
    if (-not (Test-Path "$clone\skills")) { throw "Could not fetch anthropics/skills." }

    $source = Join-Path $clone "skills"
    foreach ($skill in @("docx", "pdf", "pptx", "xlsx")) {
        $target = Join-Path $SkillsDir $skill
        if (Test-Path $target) { Remove-Item $target -Recurse -Force }
        Copy-Item (Join-Path $source $skill) $target -Recurse -Force
        Write-Host "  installed: $target"
    }

    # Windows patches: the skills ship Linux-sandbox versions that fail here.
    foreach ($skill in @("docx", "pptx", "xlsx")) {
        Copy-Item (Join-Path $root "patches\soffice.py") (Join-Path $SkillsDir "$skill\scripts\office\soffice.py") -Force
    }
    Copy-Item (Join-Path $root "patches\recalc.py") (Join-Path $SkillsDir "xlsx\scripts\recalc.py") -Force
    Write-Host "  patched: soffice.py (docx/pptx/xlsx), recalc.py (xlsx)"

    # Insert a local-environment note after the first heading of each SKILL.md.
    foreach ($skill in @("docx", "pdf", "pptx", "xlsx")) {
        $noteFile = Join-Path $root "patches\notes\$skill.md"
        $skillFile = Join-Path $SkillsDir "$skill\SKILL.md"
        if (-not (Test-Path $noteFile) -or -not (Test-Path $skillFile)) { continue }
        $note = (Get-Content $noteFile) -join "`n"
        $lines = @(Get-Content $skillFile)
        $heading = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^# ') { $heading = $i; break }
        }
        if ($heading -lt 0) { continue }
        $already = $false
        for ($i = $heading + 1; $i -lt [Math]::Min($heading + 4, $lines.Count); $i++) {
            if ($lines[$i] -match 'Local environment note') { $already = $true }
        }
        if (-not $already) {
            $new = $lines[0..$heading] + "" + $note + "" + $lines[($heading + 1)..($lines.Count - 1)]
            [System.IO.File]::WriteAllText($skillFile, ($new -join "`n"), (New-Object System.Text.UTF8Encoding($false)))
        }
    }
    Write-Host "  notes: SKILL.md local-environment notes applied"

    Write-Host "  installing Python dependencies..."
    & python -m pip install -q pypdf pdfplumber pillow python-pptx openpyxl pandas markitdown
    if ($LASTEXITCODE -ne 0) { throw "pip install of office-skill dependencies failed." }
    Write-Host "  deps: pypdf pdfplumber pillow python-pptx openpyxl pandas markitdown"
}

Write-Host "Done. Start a new agent session; the docx-agent-kit skill is now discoverable."

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
    [switch]$Link
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
Write-Host "Done. Start a new agent session; the docx-agent-kit skill is now discoverable."

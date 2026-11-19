#Requires -Version 5.1
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "docx-agent-kit installer"
Write-Host "  repo: $root"

$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    throw "Python 3.10+ is required but 'python' was not found on PATH."
}

$version = & python -c "import sys; print('.'.join(map(str, sys.version_info[:2])))"
if ([version]$version -lt [version]"3.10") {
    throw "Python 3.10+ is required (found $version)."
}

Write-Host "  python: $version"
Write-Host "  installing package (editable)..."
& python -m pip install -e $root
if ($LASTEXITCODE -ne 0) { throw "pip install failed." }

if ($env:XDG_CONFIG_HOME) {
    $config = Join-Path $env:XDG_CONFIG_HOME "opencode"
} else {
    $config = Join-Path $env:USERPROFILE ".config\opencode"
}
$skills = Join-Path $config "skills"
New-Item -ItemType Directory -Force -Path $skills | Out-Null

$dest = Join-Path $skills "docx-agent-kit"
if (Test-Path $dest) {
    $item = Get-Item $dest -Force
    if ($item.LinkType) {
        [System.IO.Directory]::Delete($dest)
    } else {
        Write-Host "  backing up existing $dest"
        Move-Item $dest "$dest.bak-$(Get-Date -Format yyyyMMddHHmmss)"
    }
}

try {
    New-Item -ItemType Junction -Path $dest -Target $root | Out-Null
    Write-Host "  skill linked: $dest"
} catch {
    Write-Host "  junction failed, copying skill files instead"
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    foreach ($name in @("SKILL.md", "README.md", "pyproject.toml", "docx_agent_kit", "tests")) {
        Copy-Item (Join-Path $root $name) -Destination $dest -Recurse -Force
    }
    Write-Host "  skill copied: $dest (re-run installer after updating the repo)"
}

& python -c "import docx_agent_kit; print('  import ok: docx_agent_kit', docx_agent_kit.__version__)"
Write-Host "Done. Start a new OpenCode session; the docx-agent-kit skill is now discoverable."

#!/usr/bin/env bash
# Install the docx-agent-kit skill into an agent's skills directory.
#
# The repository can live anywhere. This script copies the skill into the
# skills directory OpenCode or Claude Code already reads, then installs the
# Python package from that installed copy.
#
# Usage:
#   ./install.sh [--agent auto|opencode|claude] [--skills-dir PATH] [--link]
#
#   --agent       Which agent's skills directory to target (default: auto)
#   --skills-dir  Explicit directory, e.g. .opencode/skills for a project
#   --link        Link this repository instead of copying (development)

set -euo pipefail

agent="auto"
skills_dir=""
link=0

while [ $# -gt 0 ]; do
  case "$1" in
    --agent) agent="$2"; shift 2 ;;
    --skills-dir) skills_dir="$2"; shift 2 ;;
    --link) link=1; shift ;;
    -h|--help) sed -n '2,16p' "$0"; exit 0 ;;
    *) echo "Unknown option: $1" >&2; exit 1 ;;
  esac
done

root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
echo "docx-agent-kit installer"
echo "  source: $root"

PY=""
for candidate in python3 python; do
  if command -v "$candidate" >/dev/null 2>&1; then
    if "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)'; then
      PY="$candidate"
      break
    fi
  fi
done
if [ -z "$PY" ]; then
  echo "Python 3.10+ is required." >&2
  exit 1
fi
echo "  python: $($PY -c 'import sys; print(".".join(map(str, sys.version_info[:2])))')"

if [ -z "$skills_dir" ] && [ "$agent" != "auto" ]; then
  case "$agent" in
    opencode) skills_dir="${XDG_CONFIG_HOME:-$HOME/.config}/opencode/skills" ;;
    claude) skills_dir="$HOME/.claude/skills" ;;
    *) echo "Unknown agent: $agent" >&2; exit 1 ;;
  esac
fi

if [ -z "$skills_dir" ]; then
  if [ -d "${XDG_CONFIG_HOME:-$HOME/.config}/opencode" ]; then
    skills_dir="${XDG_CONFIG_HOME:-$HOME/.config}/opencode/skills"
  elif [ -d "$HOME/.claude" ]; then
    skills_dir="$HOME/.claude/skills"
  else
    skills_dir="${XDG_CONFIG_HOME:-$HOME/.config}/opencode/skills"
  fi
fi

mkdir -p "$skills_dir"
skills_dir="$(cd "$skills_dir" && pwd)"
dest="$skills_dir/docx-agent-kit"
echo "  skills: $skills_dir"

rm -rf "$dest"

install_from="$dest"
if [ "$link" -eq 1 ]; then
  ln -s "$root" "$dest"
  install_from="$root"
  echo "  linked: $dest -> $root"
else
  mkdir -p "$dest"
  for name in SKILL.md README.md pyproject.toml docx_agent_kit tests; do
    cp -R "$root/$name" "$dest/"
  done
  find "$dest" -name __pycache__ -type d -prune -exec rm -rf {} + 2>/dev/null || true
  find "$dest" -maxdepth 1 -name '*.egg-info' -type d -exec rm -rf {} + 2>/dev/null || true
  echo "  installed: $dest"
fi

echo "  installing Python package..."
"$PY" -m pip install -e "$install_from"

"$PY" -c "import docx_agent_kit; print('  import ok: docx_agent_kit', docx_agent_kit.__version__)"
echo "Done. Start a new agent session; the docx-agent-kit skill is now discoverable."

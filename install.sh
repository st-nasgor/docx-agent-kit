#!/usr/bin/env bash
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
echo "docx-agent-kit installer"
echo "  repo: $root"

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
echo "  installing package (editable)..."
"$PY" -m pip install -e "$root"

config="${XDG_CONFIG_HOME:-$HOME/.config}/opencode"
mkdir -p "$config/skills"
dest="$config/skills/docx-agent-kit"

if [ -L "$dest" ]; then
  rm "$dest"
elif [ -e "$dest" ]; then
  backup="$dest.bak-$(date +%Y%m%d%H%M%S)"
  echo "  backing up existing $dest"
  mv "$dest" "$backup"
fi

ln -s "$root" "$dest"
echo "  skill linked: $dest"

"$PY" -c "import docx_agent_kit; print('  import ok: docx_agent_kit', docx_agent_kit.__version__)"
echo "Done. Start a new OpenCode session; the docx-agent-kit skill is now discoverable."

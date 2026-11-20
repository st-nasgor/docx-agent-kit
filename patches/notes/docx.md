> **Local environment note.** Node and pandoc are typically not installed on
> this setup. For creating documents use Python 3.10+ with `python-docx` plus
> the `docx-agent-kit` pack (`import docx_agent_kit`,
> https://github.com/st-nasgor/docx-agent-kit), which provides hyperlinks,
> fields, bullets/numbering, table borders/shading, image swaps, template
> placeholders and clean-up — do not hand-roll that OOXML. For reading, use
> `python-docx` or `python -m docx_agent_kit inspect` instead of pandoc. The
> validator, `soffice.py`, `merge_runs.py` and `accept_changes.py` in this
> skill work with the installed Python and LibreOffice; `docx-agent-kit` also
> exposes their common operations. LibreOffice CLI: `soffice.com` on Windows
> (`%ProgramFiles%\LibreOffice\program\soffice.com`), `soffice` on macOS/Linux.

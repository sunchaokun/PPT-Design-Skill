# Installation and runtime

## Python package

The skill depends on the published `pptx-designer` package:

```powershell
python -m pip install --upgrade pptx-designer
```

Core rendering dependencies are installed with the package. Optional image
search or generation features may require:

```powershell
python -m pip install --upgrade "pptx-designer[images]"
python -m pip install --upgrade "pptx-designer[ai-images]"
```

Use `skill/scripts/check_runtime.py` before a generation task. It reports
missing Python or rendering dependencies without silently changing the user's
environment.

## PPTX → PDF → PNG dependencies

The preferred Windows renderer is Microsoft PowerPoint through COM. The
headless fallback requires:

- LibreOffice (`soffice` or `soffice.bin`)
- Poppler (`pdftoppm`)

These are operating-system dependencies, not Python dependencies. Install
them through the user's approved system package mechanism and then rerun the
runtime check. The repository installer offers an explicit Windows-only
`python install.py --render-deps` option using `winget`; it is never enabled
implicitly.

## Fonts

Fonts are part of visual correctness. Detect the target language and use fonts
available on the target machine. If a requested font is unavailable, report
the substitution before final delivery and inspect the PNG again.

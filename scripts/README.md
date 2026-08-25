# Compatibility scripts

The old repository exposed rendering and inspection commands from `scripts/`.
The maintained implementations now live under `skill/scripts/` because they
belong to the installed skill bundle.

Use:

```powershell
python skill/scripts/check_runtime.py
python skill/scripts/inspect_pptx.py path/to/deck.pptx --pretty
powershell -ExecutionPolicy Bypass -File skill/scripts/render_pptx.ps1 `
  -InFile path/to/deck.pptx -OutDir output/rendered
```

The former engine-specific `generate_ppt.py`, `audit_pptx.py`, and
`layout_report.py` are intentionally not copied. Their implementation was
tied to `ppt_pro_max`; generation now belongs to `pptx-designer`, while final
visual acceptance belongs to LLM PNG review.

# PPT Design Skill

This repository contains the `ppt-design-skill` workflow. The PowerPoint
generation engine is the published `pptx-designer` package; do not recreate or
vendor the former `ppt_pro_max` implementation here.

## Source of truth

- `skill/SKILL.md`: the agent workflow and design rules;
- `skill/references/`: public API, domain paradigms, content rules, and QA;
- `skill/scripts/`: runtime, structural inspection, and PPTX-PDF-PNG rendering;
- `examples/output/`: three real editorial PPTX reference cases;
- `installer/`: cross-platform skill installation;
- `README.md` and `docs/`: user-facing documentation.

## Required generation loop

1. Confirm the brief, page structure, and brief acceptance contract.
2. Choose FreeStyle `generate_ppt()`, Build Mode, or VI Build Mode.
3. Generate a reproducible PPTX.
4. Reopen it with `skill/scripts/inspect_pptx.py`.
5. Render through PPTX -> PDF -> PNG.
6. Compare every PNG with the acceptance contract, record evidence, and revise
   material visual issues or failed `MUST` requirements.

Do not claim completion from a successful Python run alone.

## Validation

```powershell
python skill/scripts/check_runtime.py
python installer/install.py --check
python skill/scripts/inspect_pptx.py examples/output/luxury_fragrance_lookbook.pptx --pretty
python "C:\Users\Administrator\.codex\skills\.system\skill-creator\scripts\quick_validate.py" skill
```

## Naming and compatibility

The skill name is permanently `ppt-design-skill`. Do not rename installation
folders to `pptx-designer`; that is the Python package name, not the skill name.

## Editing rules

- Use documented public `pptx_designer` APIs in examples and generated code.
- Keep the three reference PPTX artifacts intact unless the user asks for a
  replacement.
- Preserve the early skill's design rules when revising documentation.
- Use `apply_patch` for text and code edits.

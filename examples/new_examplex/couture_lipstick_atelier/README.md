# COUTURE COLOR — Objects of Desire

A 10-slide English editorial portfolio case for a fictional couture lipstick
collection. It demonstrates a premium product story without borrowing a real
brand or presenting fabricated commercial claims.

## Deliverables

- `build.py` — reproducible Build Mode source.
- `assets/images/` — five original AI-generated editorial photographs,
  including two frames of the same model in a product-in-use beauty sequence.
- `output/couture_color_objects_of_desire.pptx` — editable presentation.
- `rendered/couture_color_objects_of_desire.pdf` and `rendered/` PNGs — final
  reviewed export.
- `brief.md`, `page-plan.md`, `visual-direction.md`, and
  `acceptance-contract.md` — design rationale and verification scope.

## Rebuild

```powershell
python build.py
powershell -ExecutionPolicy Bypass -File "E:\PPT-Design-Skill-V2\skill\scripts\render_pptx.ps1" `
  -InFile output\couture_color_objects_of_desire.pptx `
  -OutDir rendered
```

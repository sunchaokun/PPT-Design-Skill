# Acceptance contract

| ID | Level | Requirement | Expected evidence in rendered output | Status |
|---|---|---|---|---|
| R1 | MUST | FreeStyle accepts a validated resolved theme without discovery arguments. | `freestyle-render/slide01.png` and `slide02.png` show coherent dark-tech output; `ignored_arguments` is empty. | PASS |
| R2 | MUST | Build Mode accepts the same resolved theme with `strict_theme=True`. | `build-render/slide01.png` and `slide02.png` show 30 native text/shape objects across two slides. | PASS |
| R3 | MUST | Build content is readable and unclipped at presentation scale. | All four 1280×720 PNGs were visually reviewed; no clipping, overlap, or missing text observed. | PASS |
| R4 | MUST | VI protected merge retains a template lock and reports an attempted lock override. | `output/contract-results.json` records an empty allowed conflict list and a rejected `assets.logo` conflict. | PASS |

Rules:

- Do not rewrite a MUST requirement merely to fit an output.
- Every MUST requires visible evidence in the rendered pages.
- Record `PASS`, `NEEDS_REVISION`, or `BLOCKED` with evidence.

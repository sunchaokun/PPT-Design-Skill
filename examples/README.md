# Presentation Case Studies

This directory contains the maintained case-study library for PPT Design Skill.
Each case is a complete, editable PowerPoint project rather than a decorative
mock-up: it pairs a delivered PPTX and PDF with the source material, build
logic, page plan, visual direction, and review evidence used to produce it.

The browser gallery is in [site/](site/). Its published case list, preview
images, and downloadable files are defined in
[site/data/examples.json](site/data/examples.json).

## Maintained cases

| Case | Pages | Focus | Project package |
|---|---:|---|---|
| AI Agent Operating System | 12 | Technical architecture, control planes, evaluation, and safety | [Open](new_examplex/ai_agent_operating_system/) |
| AI Infrastructure Economics | 12 | Capital, compute, energy, platform constraints, and operating choices | [Open](new_examplex/ai_infrastructure_economics/) |
| Single-Cell CAR T Atlas | 12 | Paper-led scientific narrative, evidence boundaries, and figure-based explanation | [Open](new_examplex/car_t_single_cell_paper/) |
| Louvre Abu Dhabi | 10 | Architectural atmosphere, geometry, climate, and cultural space | [Open](new_examplex/louvre_abudhabi/) |
| Vertical City Retrofit | 14 | Urban retrofit strategy, systems thinking, phasing, and investment decisions | [Open](new_examplex/vertical_city_retrofit/) |
| COUTURE COLOR — Objects of Desire | 10 | Luxury beauty editorial, product-in-use imagery, editable product geometry, and material storytelling | [Open](new_examplex/couture_lipstick_atelier/) |

Together, these six projects provide seventy reviewed slides across technical,
data, scientific, cultural, urban-strategy, and luxury-beauty presentation domains.

## What a complete case includes

The exact file set differs by project, but a maintained case normally contains:

- `brief.md` — audience, scenario, content boundary, and intended takeaway.
- `visual-direction.md` — the visual thesis, typography, palette, grid, image
  treatment, and prohibited patterns.
- `page-plan.md` — the role and visual form of every slide.
- `acceptance-contract.md` and `visual-review.md` — the requirements and
  evidence from the final review.
- `build.py` or equivalent Build Mode source — the reproducible authoring
  entry point.
- `output/` — the delivered editable `.pptx` and exported `.pdf`.
- `rendered/` — reviewed PNG slides when the case retains them.

Cases that use third-party research figures or photography include source,
credit, or usage-boundary notes where applicable. Those materials must be
reviewed before any external commercial reuse.

## Browse or inspect a case

Open the local gallery from [site/index.html](site/index.html), or serve the
repository through a static web server. Each gallery entry opens a page-level
viewer and links to the matching PPTX and PDF.

To inspect a delivered deck structurally, run the project command from the
repository root. For example:

```powershell
python skill/scripts/inspect_pptx.py `
  examples/new_examplex/louvre_abudhabi/output/louvre_abudhabi_complete.pptx `
  --pretty
```

For a full visual check, render the PPTX to PDF and PNG, then inspect every
slide at presentation scale:

```powershell
powershell -ExecutionPolicy Bypass -File skill/scripts/render_pptx.ps1 `
  -InFile examples/new_examplex/louvre_abudhabi/output/louvre_abudhabi_complete.pptx `
  -OutDir output/louvre_abudhabi_review
```

Do not treat a successful build as final acceptance. A case is ready only when
its rendered slides satisfy the brief, visual direction, and acceptance
contract.

## Contribution standard

Add a new case only when it expands the library with a distinct communication
problem or visual language. It must include a reproducible source, an editable
PPTX, a rendered review pass, and clear provenance for any external assets.
Avoid adding partial drafts, duplicate exports, or one-off assets to the
gallery.

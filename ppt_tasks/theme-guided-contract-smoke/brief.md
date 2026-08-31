# Theme-guided contract smoke test

- Topic: Verify theme-guided generation contracts in `ppt-design-skill`.
- Audience: Skill maintainers and presentation-engine developers.
- Scenario: Internal release validation for `pptx-designer` 1.0.0b9.
- Language: English.
- Purpose / decision / action: Confirm the documented FreeStyle, Build, and VI
  theme paths are usable and yield editable, rendered PPTX output.
- Duration and page count: Two Build slides and two FreeStyle slides.
- Source material: Resolved theme generated locally from `dark-tech`, seed 17.
- Brand, image, data, and editability constraints: No external images; native
  editable text and shapes only.

## Outcome

The maintainer can confirm that one resolved theme is accepted by both
generation modes, its output survives PNG rendering, and VI template locks are
protected rather than silently overridden.

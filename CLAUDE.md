# PPT Design Skill

Use the `ppt-design-skill` workflow for presentation design tasks. The
published `pptx-designer` Python library creates the editable PPTX; this
repository owns the design process, installation, rendering, and PNG review.

Before changing the skill, read:

- `skill/SKILL.md`
- `skill/references/design-principles.md`
- `skill/references/qa-and-delivery.md`
- `docs/README.md`

The final acceptance gate is PPTX -> PDF -> PNG followed by direct LLM review
of the rendered PNGs. Keep the skill name `ppt-design-skill` in all metadata
and installer paths.

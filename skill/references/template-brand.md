# VI Build and template brand workflow

VI Build is the third skill mode. Use it when a user supplies an existing
PowerPoint template or explicitly requires enterprise brand compliance.

## Mode decision

- No template and exact composition required: Build Mode.
- No template and speed or goal-driven generation is acceptable: FreeStyle.
- A template or corporate master must be preserved: VI Build.

## Workflow

1. Reopen the template and inspect page size, slide count, text zones, fonts,
   colors, logos, recurring decorations, and master assumptions.
2. Call `extract_design_dna(template_path)` where the public package supports
   the required analysis.
3. Extract a template context and record its locked fields: background,
   primary/accent colors, heading/body fonts, safe margins, logo placement,
   footer rules, and allowed components.
4. Keep framework pages such as cover, agenda, section pages, and closing page
   unchanged unless the user explicitly asks for redesign.
5. Resolve any approved Theme Lock, then call
   `merge_vi_design_context(template_context, resolved_theme, page_context)`.
   Do not use generic `merge_design_context()` for this policy boundary:
   later inputs may overwrite it. Stop when `diagnostics.conflicts` is not
   empty. To permit an exceptional brand change, first create a newly approved
   template context with that lock deliberately changed or removed, then rerun
   the protected merge; the rejected context must never be passed through.
6. Add new pages using the template as the starting presentation and public
   `pptx_designer` components.
7. Render the complete result through PPTX -> PDF -> PNG and check both
   preserved and newly added pages.

## Page-context example

```python
page_context = {
    "visual_grammar": {
        "allowed_atom_kinds": ["text", "shape", "chart"],
    },
    "acceptance": {
        "must_coverage": ["brand_footer_present"],
    },
}
```

The page context constrains the new pages; it is not permission to replace the
template with a blank presentation or to overwrite template locks. If master
behavior cannot be preserved through the public API, report the limitation and
ask whether approximation is acceptable.

## Acceptance criteria

- framework pages remain intact;
- logo and recurring brand elements are not duplicated or misplaced;
- new pages use the same margins, fonts, color roles, and footer language;
- no page introduces an unrelated palette or component family;
- all pages pass PNG visual review;
- editable content remains editable wherever supported.

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
3. Record a VI token containing background, primary/accent colors, heading/body
   fonts, safe margins, logo placement, footer rules, and allowed components.
4. Keep framework pages such as cover, agenda, section pages, and closing page
   unchanged unless the user explicitly asks for redesign.
5. Add new pages using the template as the starting presentation and public
   `pptx_designer` components.
6. Render the complete result through PPTX -> PDF -> PNG and check both
   preserved and newly added pages.

## Token example

```python
VI = {
    "primary": "#1E3A5F",
    "accent": "#C9A96E",
    "background": "#F8FAFC",
    "text_dark": "#1A2B3C",
    "text_body": "#37474F",
    "font_heading": "Aptos Display",
    "font_body": "Aptos",
    "safe_margin": 0.65,
}
```

The token constrains the new pages; it is not permission to replace the
template with a blank presentation. If master behavior cannot be preserved
through the public API, report the limitation and ask whether approximation is
acceptable.

## Acceptance criteria

- framework pages remain intact;
- logo and recurring brand elements are not duplicated or misplaced;
- new pages use the same margins, fonts, color roles, and footer language;
- no page introduces an unrelated palette or component family;
- all pages pass PNG visual review;
- editable content remains editable wherever supported.

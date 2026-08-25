# Structured content mode

Use `generate_ppt(content={...})` when the page structure is known but exact
element placement is not required.

```python
content = {
    "title": "Q4 Revenue Report",
    "pages": [
        {
            "goal": "hook",
            "title": "Q4 2026",
            "subtitle": "Record quarter",
        },
        {
            "goal": "content",
            "title": "What changed",
            "bullets": ["Revenue grew 23%", "Enterprise mix increased"],
        },
        {
            "goal": "data",
            "title": "Key metrics",
            "bullets": ["Revenue: $12.8M", "Retention: 89%"],
        },
    ],
}
```

Use concise bullets. The `goal` should describe the communication role, not
the visual decoration. Typical goals include `hook`, `problem`, `solution`,
`features`, `data`, `code`, `exercise`, `overview`, `content`, and `cta`.

When a page needs custom diagrams, exact coordinates, advanced composition, or
domain-specific figure layout, switch to Build Mode instead of forcing the
content schema to express it.

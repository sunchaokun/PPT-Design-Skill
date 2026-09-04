# Object map

The source of truth is `../build.py` and the generated PPTX at
`../output/louvre_abudhabi_complete.pptx`. The map covers P01, P03, and P07.

| Object | Implementation | Editable | Evidence / boundary |
|---|---|---:|---|
| Titles, labels, captions, source notes | native text | yes | inspect_pptx reports text boxes |
| Hero and supporting photographs | native picture | yes | crop can be changed in PowerPoint |
| Dome / star geometry | native freeform | yes | inspect_pptx reports FREEFORM objects |
| Technical labels and dimensions | native text / shapes | yes | factual content remains replaceable |
| Optional local contour treatment | SVG-assisted only | partial | decorative/local geometry, never text or data |
| Full-page composition | native mixed objects | yes | no full-page screenshot source |

# Design direction coach

The user may know the topic but not the language of presentation design. The
skill must act as a design advisor, not ask the user to choose unexplained
parameters such as `variance`, `density`, or “what style do you want?”.

## 1. Translate the brief before asking for style

Infer an initial design read from:

- audience and decision context;
- domain and cultural conventions;
- presentation occasion and delivery surface;
- content type and evidence burden;
- desired audience response;
- brand, template, image, and editability constraints.

Ask only the missing questions that could materially change the design. Explain
why each question matters in plain language.

## 2. Recommend directions, do not outsource design thinking

For a delivery-grade task whose visual direction is not fixed, present two or
three genuinely different directions. Each direction should include:

1. a memorable name;
2. a one-sentence visual thesis;
3. audience and occasion fit;
4. palette, typography, grid, image, and decoration language;
5. page archetypes and narrative rhythm;
6. what it intentionally avoids;
7. its tradeoff or risk;
8. a clear recommendation.

Do not present “blue / green / purple” as separate directions. A direction is
different only when its visual grammar and page architecture are different.

Example for a couture brief:

```text
A. Atelier Research / recommended
Concept: a study of volume, material, and the body.
Grammar: asymmetric editorial grid, bone white and ink, serif display type,
material labels, silhouette crops, large quiet fields.
Avoids: product catalog cards and symmetric feature rows.

B. Nocturnal Runway
Concept: movement and reveal under controlled light.
Grammar: near-black ground, cinematic crops, oversized type, sharp accent line,
sequenced reveals.
Risk: can become too campaign-like if the content is evidence-heavy.

C. Material Archive
Concept: a curatorial index of craft and provenance.
Grammar: specimen labels, numbered studies, texture panels, restrained grid,
captions and catalog rhythm.
Risk: more archival and less emotionally immediate.
```

## 3. Make novice confirmation easy

Use a short confirmation card:

```text
推荐方向：A. Atelier Research
它适合：高定概念发布 / 私人 editorial / 设计研究
它会让观众感到：克制、专业、具有艺术指导感
它不会做成：普通产品画册或三栏卖点页
主要代价：需要更高质量的图片和更精炼的文案

请确认：采用 A，或告诉我你更接近 B/C 的哪一部分。
```

If the user cannot decide, choose the recommended direction and state the
assumption. Do not block the task with an abstract design questionnaire.

## 4. Lock the visual direction

After confirmation, write a visual-direction lock containing:

- visual thesis and audience promise;
- must-have visual signatures;
- palette and typography roles;
- grid, margins, spacing, and image treatment;
- page archetype map and density rhythm;
- forbidden patterns and fallback rules;
- the acceptance-contract rows that prove the direction survived in PNG.

The lock is the source of truth for Build, FreeStyle, and VI Build. A PNG
review must check not only whether pages are readable, but whether the locked
visual thesis is still visible in the final result.

## 5. Call the library at two different moments

The built-in `pptx-designer` design data should be consulted twice, with two
different responsibilities:

### A. Before direction lock: discovery and recommendation

Use the library to widen and ground the design conversation:

```python
from pptx_designer import recommend_styles
from pptx_designer.renderer.theme import ThemeComposer
from pptx_designer.search.adapters import search_color, search_style, search_typography

candidates = recommend_styles("couture fashion editorial", top_k=3)
theme_hint = ThemeComposer().compose(query="couture fashion editorial", seed=7)
font_options = search_typography("serif", top_k=3)
```

Use the results as candidate atoms and vocabulary for two or three directions.
Filter them through the audience, domain, visual thesis, and forbidden-pattern
rules. The library may return a generic luxury or corporate preset for an
ambiguous query; the LLM must not present that result as the final art
direction without interpretation.

### B. After direction lock: implementation and reproducibility

Once the user selects a direction, pin explicit choices rather than repeatedly
sampling styles:

```python
from pptx_designer.renderer.theme import ThemeComposer

theme = ThemeComposer().compose(
    palette="golden-luxury",
    fonts="serif-editorial",
    decoration="gold-trim",
    layout="asymmetric",
    seed=7,
)
```

For Build Mode, translate the selected theme into an explicit `C` dictionary,
typography roles, image rules, and page archetypes. For FreeStyle, pass pinned
`palette`, `fonts`, `decoration`, `layout`, `mood`, and `style_seed` values when
the public function supports them. The user-confirmed visual lock remains the
authority; library atoms are implementation inputs, not a replacement for
design judgment.

Do not call the library to randomly change direction after every PNG review.
Use PNG review to diagnose the visible result against the lock, then make a
targeted change to the source or pinned theme.

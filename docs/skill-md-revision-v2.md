# SKILL.md 修订方案 v2 — 5 问题修正版

基于 `skill-md-context-budget-review.md` 的审查结论，修正 5 个内部矛盾后重新设计。

## 问题 1：AI Tells 去重

### 重复项

| 重复组 | 项 A | 项 B | 修正 |
|--------|------|------|------|
| 状态点 | "禁止装饰性彩色状态点" (line 197) | "禁止装饰性状态点" (line 219) | 合并为 1 条：**禁止装饰性状态点**（彩色是子集，总集更强） |
| 地域时间 | "禁止天气/城市/时间条" (line 206) | "禁止城市/时间/天气条" (line 217) | 合并为 1 条：**禁止城市/时间/天气条**（同一条，仅排序不同） |

### 去重后的 AI Tells 列表（25 条，原 28 条删 3 条）

```
### AI Tells Blacklist (HARD BAN unless user explicitly requests)

- No cover version labels (V0.6/BETA/内测版)
- No "Brand · No.01" style sub-labels
- No section-number eyebrows (00/INDEX, 001·核心能力) — use natural language
- No card/image numbering labels (01/04, 1/3)
- Max 1 middle-dot (·) per metadata line — no "foo · bar · baz · qux"
- No decorative status dots (colored or uncolored)
- No em-dash (—) or Chinese em-dash (——) — use comma/hyphen/semicolon
- No linebreak+italic "design trick" (for thirty\n*years.*)
- No vertical rotated text
- No crosshair/grid-line decoration (only for organizing real content)
- No fake product UI (text-box dashboard/terminal/task-list)
- No fake version footers (v0.6.2-rc.1, "last sync 4s ago")
- No "silently used by" / "默默服务" social-proof headlines — use natural language or skip
- No "来自一线" / "实战笔记" artisan labels — use plain functional labels
- No city/time/weather bars (99% of scenarios)
- No eyebrow micro-metadata sentences
- No generic step labels ("Phase 1/2/3", "步骤 01/02/03") — use verb+noun ("Install", "Configure")
- No overlaid labels on images ("Brand · 02", "场景 - Demo")
- No decorative photo credits (场景 III · 35mm) — skip or use one-line caption
- No version footers on marketing slides (v1.4.2, Build 0048)
- No inventory counters as decoration ("已预约 412/800")
- No bottom-of-cover decoration strips (品牌. 创新. 技术.)
- No floating explanation text top-right of section titles
- No divider lines on every row of long lists
- No progress bars with filled background tracks for comparison visuals
- No scroll hints (Scroll, ↓向下滚动)
```

**净减：3 条**（2 组去重 + "禁止滚动提示"和"禁止天气条"合并到"禁止城市/时间/天气条"）

## 问题 2：Content Design Rules 与 Design Constraints 合并

### 冲突分析

| 现有 Content Design Rules | 新增 Design Constraints | 冲突类型 |
|---------------------------|------------------------|----------|
| "features: first card is the most important, with longer body text" | "features 页首张卡片必须 featured" | **重复** — 说同一件事 |
| "6+ bullets trigger two-column layout" | "6+双栏" | **重复** — 说同一件事 |
| "vary bullet density across pages" | "连续3页禁止都是标题+要点" | **互补** — 现有是正面建议，新增是负面约束 |
| "use concrete numbers and data" | "禁止假精确数字" | **互补** — 鼓励真数据 vs 禁止假数据 |

### 合并方案：Content Design Rules 升级

**删除 Design Constraints 中与 Content Design Rules 重复的条目，将 Design Constraints 的增量信息合并进 Content Design Rules 表格。**

原 Content Design Rules 是 8 行表格，升级为 12 行（+4 行增量约束）：

```markdown
### Content Design Rules (CRITICAL — maximizes design quality)

| Rule | Why | Example |
|------|-----|---------|
| features: first card featured with longer body | First card gets gradient bar + 22pt title + higher elevation | Card 1: "智能推理引擎 — 自动选择最优框架" vs Card 2: "全链路监控" |
| 6+ bullets → two-column layout | Better density; layout engine auto-splits | 6 concise data points instead of 3 long ones |
| tech topics: include code page | Code pages add technical credibility | `{"code": {"language": "python", "source": "..."}}` |
| education/training: include exercise page | Exercise pages add interactivity | `{"exercise": {"duration": "5 min", "steps": [...]}}` |
| topic transitions: insert section divider | Visual rhythm (oversized number + gradient line) | Between problem→solution |
| hook: short subtitle (<40 chars); cta: long (>60) | Different hero compositions | hook: "5分钟取代5周" vs cta: "免费额度包含1000次..." |
| vary bullet density (some 3-bullet, some 6+) | Varying density feels natural; 10+ items → cards/grid/table, never list | Don't make every page the same density |
| use concrete real data; no fake precision | "GPU成本年增3倍" not "成本持续增长"; no fabricated 92%/4.1× | Real data only; mark as "example" if hypothetical |
| ≤5 bullets: single column | 6+: two-column | 10+: use cards/grid/infographic component, never list |
| no filler verbs (赋能/领先/一站式/生态/革新/引领) | AI-generated buzzwords destroy credibility | Use plain functional language |
| quotes ≤3 lines, attribution = name+title | PPT quotes are fragments, not full reviews | "Name, CTO, Company" — never name alone |
| theme lock: one theme per deck, no mid-deck switch | Dark stays dark, light stays light; micro-variation OK | #0A1E3D → #0F2847 OK; #0A1E3D → #FFF8F0 NOT OK |
```

**Design Constraints 中删除的条目**（已合并到 Content Design Rules）：
- "features 页首张卡片必须 featured" → 已在 Rule 1
- "6+双栏" → 已在 Rule 2+9
- "禁止假精确数字" → 已在 Rule 8
- "禁止填充动词" → 已在 Rule 10
- "引文≤3行，署名=姓名+职位" → 已在 Rule 11
- "主题锁定" → 已在 Rule 12

**Design Constraints 中保留的条目**（Content Design Rules 未覆盖的）：
- ≤5 单列 / 10+ 不用列表 → Rule 9 新增，但"禁止20+行数据表"仍需在 Layout 约束中
- "禁止规格页用10行要点" → 属于 Layout 约束
- "每页：短标题+短副文+1视觉/CTA" → 属于 Typography 约束

## 问题 3：中英文统一

### 原则

- **SKILL.md 全英文**（与现有 634 行一致）
- **中文术语只在括号内标注**（帮助中文场景的 LLM 理解语境）
- **Design Vocabulary 中的模式名必须映射到 content.json 的 goal/layout 参数**

### 修正方案

| 原中文 | 英文 | content.json 映射 |
|--------|------|-------------------|
| 呼吸页 | Breathing slide (low-density) | `goal: "content"`, ≤2 bullets, large margins |
| 过渡页 | Section divider | `goal: "section"` |
| 数据冲击页 | Data-impact slide | `goal: "data"`, single big number |
| 视觉锚点页 | Visual anchor (full-bleed image) | `goal: "content"`, image with minimal text |
| 行动页 | CTA / Action slide | `goal: "cta"` |
| 强调色 | Accent color | brand.json `accent` |
| 一致性锁 | Consistency lock | — |
| 填充动词 | Filler verbs | — |
| 假精确数字 | Fake precision numbers | — |
| 主题锁定 | Theme lock | — |

### Design Constraints 全文英文化示例

```markdown
### Layout
- VARIANCE > 4: avoid all-centered; use left-aligned / sidebar / asymmetric
- features page: first card MUST be featured (larger title + gradient bar)
- Cover title ≤2 lines, subtitle ≤20 chars, hero ≤4 text elements
- Same layout family max 1 occurrence per deck
- 8+ slides: ≥4 distinct visual layouts
- Left-right alternation ≤2 times; 3rd time = break the pattern
- Eyebrow count ≤ceil(page_count / 3); NO section-number eyebrows
- No split-header as default (title left + explanation right)
- Bento grid: ≥2-3 cells with visual variation (not all-white text cards)
```

## 问题 4：Dial → LLM Action Mapping

### 问题本质

Step 1 要求 LLM 声明 V/M/D，但没告诉 LLM "声明后写 content.json 时该做什么不同"。没有动作映射，Dial 是空声明。

### 解决方案：Dial Action Table

在 Step 1 的 Dial 声明后，直接给出每个 Dial 值范围对应的 **LLM 动作**（不是渲染行为，是 LLM 在 content.json 中的决策行为）：

```markdown
**Dial → content.json Action Map:**

| Dial Range | LLM Action in content.json |
|-------------|---------------------------|
| VARIANCE 1-3 | Use centered layouts; equal-width cards; symmetric spacing; `layout_variant: "centered"` or `"standard"` |
| VARIANCE 4-7 | Left-align titles; feature first card; use `layout_variant: "sidebar-left"` or `"asymmetric"`; vary image position |
| VARIANCE 8-10 | Use unequal grids; asymmetric hero; mix `layout_variant` across slides; avoid any repeated layout family |
| MOTION 1-3 | No animation entries in content.json; rely on slide transitions only |
| MOTION 4-7 | Add `animation` hints on key slides (cover, section dividers); use fade-in/stagger pattern |
| MOTION 8-10 | Add `animation` hints on most slides; fly-in entrances; sequence reveals; grow on CTA |
| DENSITY 1-3 | 2-3 bullets per slide; generous spacing; `density: 3` parameter; breathing slides after every content slide |
| DENSITY 4-7 | 3-5 bullets per slide; standard spacing; `density: 5`; mix densities across slides |
| DENSITY 8-10 | 6+ bullets on data slides; compact tables; `density: 8`; infographic components; no breathing slides |
```

**关键设计决策**：Dial 的动作映射是 **LLM 行为级的**，不是渲染器 API 级的。LLM 根据 Dial 值选择不同的 content.json 写法（goal type、bullet 数量、layout_variant、animation hint），而不是直接调用渲染器参数。这解决了"Dial 是声明式不执行"的问题 — LLM 自己执行 Dial 的意图。

### content.json 新增字段

为支持 Dial 动作映射，content.json 需要两个可选字段：

```json
{
  "goal": "content",
  "title": "...",
  "layout_variant": "sidebar-left",
  "animation": "fade-in",
  ...
}
```

- `layout_variant`: 可选，覆盖 DesignDecider 的自动选择。值 = Design Atoms 中的 8 种变体
- `animation`: 可选，指定该页的入场动效。值 = Animation System 中的 11 种

**注意**：这两个字段是 LLM 根据 Dial 推理主动写入的，不是系统自动注入的。如果 LLM 不写，渲染器用默认值（行为不变）。

## 问题 5：Consistency Locks 与 Color 去重

### 重复项

| Color 节 | Consistency Locks 节 | 修正 |
|----------|---------------------|------|
| "全PPT ≤1个强调色，贯穿每一页" | "强调色贯穿全PPT" | 删 Locks 中的重复，Color 节保留完整版 |
| "冷暖灰不混用 — 全冷 or 全暖" | "冷暖灰不混用" | 删 Locks 中的重复，Color 节保留完整版 |
| — | "字体对一致（标题+正文，全篇统一）" | 保留在 Locks（Color 节不含此项） |
| — | "圆角体系一致" | 保留在 Locks（Color 节不含此项） |
| — | "暗色/亮色主题不中途切换" | 保留在 Locks（Content Design Rules Rule 12 已含 → 删 Locks 中的） |
| "暗色主题文字对比度≥60%亮度差" | — | 仅在 Color 节 |

### 修正后的结构

**Consistency Locks 节删除 3 条重复**，只保留 Color 和 Content Design Rules 未覆盖的一致性约束：

```markdown
### Consistency Locks
- Corner radius: pick ONE system for entire deck — sharp (0pt) / soft (8-12pt) / pill — no mixing
- Font pair: consistent throughout (heading font + body font, unified across all slides)
- (Accent color, warm/cool gray, and theme lock are already enforced in Color rules and Content Design Rules)
```

从 5 条减到 2 条 + 1 行交叉引用。

---

## 修正后的完整 Design Constraints 结构（英文版）

```markdown
## Design Constraints

### Typography
- Cover title: 44-52pt, ≤2 lines | Inner title: 32-36pt | Body: 14-16pt | Bullets: 13-14pt | Caption: 11-12pt
- Min 4 font-size levels per deck — 2 levels (title+body only) is forbidden
- NO Calibri/Arial as default font (PPT's AI default, same as Inter in web)
- Serif: only for editorial / luxury / heritage scenes — NOT for tech/startup/data
- Max 2 font families per deck (heading + body, +1 monospace for code)
- Emphasize with bold or color shift — NO mixing serif+sans for "contrast emphasis"
- Italic title line-height ≥1.1× (descender clearance)
- Each page: short title (≤8 words) + short subtitle (≤25 chars) + 1 visual OR 1 CTA

### Color
- Max 1 accent color per deck, used on EVERY page (consistency lock)
- NO default-blue gradient cover when style is unspecified
- NO default gold+navy for "premium" scenes (#1A1A2E / #C9A96E family)
- Warm/cool gray: pick one, use throughout — no mixing
- Chart colors derived from main palette — no rainbow
- Dark theme: text ≥60% luminance above background (projection-grade contrast)
- Light theme: no light-gray text on white (invisible on projector)

### Layout
- VARIANCE > 4: avoid all-centered; use left-aligned / sidebar / asymmetric
- Same layout family max 1 occurrence per deck ("Our Products" ≠ "Core Advantages" visually)
- 8+ slides: ≥4 distinct visual layouts
- Left-right alternation ≤2 times; 3rd = break pattern
- Eyebrow count ≤ceil(page_count / 3); NO section-number eyebrows (00/INDEX, 001·核心能力)
- No split-header as default (title left + small text right)
- Bento grid: ≥2-3 cells with visual variation (not all-white text cards)
- Spec sheets: NO 10-line bullet lists — use card grid / highlight+fold / grouped sections
- NO 20+ row data tables in PPT — PPT is presentation, not document

### Page Roles
- MUST plan 5 roles: breathing (low-density rest) / section-divider / data-impact (big number) / visual-anchor (full-bleed image) / cta
- No 3 consecutive "title + bullets" pages
- 6+ slides: ≥1 section divider
- High-density page → must be followed by low-density page

### Visual Assets
- Cover MUST have real visual (not text + gradient block)
- Even minimalist style: ≥2-3 pages with images
- NO fake product screenshots (text-box dashboards/terminals/task-lists — #1 AI Tell in PPT)
- Logo wall: use real logo images, not text spans

### Consistency Locks
- Corner radius: ONE system per deck — sharp (0pt) / soft (8-12pt) / pill — no mixing
- Font pair: consistent throughout (heading + body, all slides)
- Accent color + warm/cool gray + theme lock: enforced in Color rules and Content Design Rules

### AI Tells Blacklist (HARD BAN unless user explicitly requests)

- No cover version labels (V0.6/BETA/内测版)
- No "Brand · No.01" style sub-labels
- No section-number eyebrows (00/INDEX, 001·核心能力) — use natural language
- No card/image numbering labels (01/04, 1/3)
- Max 1 middle-dot (·) per metadata line — no "foo · bar · baz · qux"
- No decorative status dots
- No em-dash (—) or Chinese em-dash (——) — use comma/hyphen/semicolon
- No linebreak+italic "design trick"
- No vertical rotated text
- No crosshair/fine-grid decoration (only for organizing real content)
- No fake product UI (text-box dashboard/terminal/task-list)
- No fake version footers (v0.6.2-rc.1, "last sync 4s ago")
- No "silently used by" / "默默服务" social-proof headlines — use natural language or skip
- No "来自一线" / "实战笔记" artisan labels — use plain functional labels
- No city/time/weather bars (99% of scenarios)
- No eyebrow micro-metadata sentences
- No generic step labels ("Phase 1/2/3", "步骤 01/02/03") — use verb+noun
- No overlaid labels on images ("Brand · 02")
- No decorative photo credits (场景 III · 35mm) — skip or use one-line caption
- No version footers on marketing slides (v1.4.2, Build 0048)
- No inventory counters as decoration ("已预约 412/800")
- No bottom-of-cover decoration strips (品牌. 创新. 技术.)
- No floating explanation text top-right of section titles
- No divider lines on every row of long lists
- No progress bars with filled background tracks for comparison
- No scroll hints (Scroll, ↓)

### Design Vocabulary (pattern → trigger condition → content.json mapping)

**Covers**: Asymmetric Split Hero → `goal:"hook"`, image on one side | Editorial Manifesto → `goal:"hook"`, no image, large type | Full-Bleed Image → `goal:"hook"`, image with overlay | Data-Impact → `goal:"hook"`, big number + one-liner | Minimal Typography → `goal:"hook"`, text-only, extreme whitespace

**Inner pages**: Sidebar+Content → `layout_variant:"sidebar-left"`, consulting/reports | Split Text-Image → `goal:"content"`, image field | Bento Grid → `goal:"features"`, 4+ cards | Big Number Focus → `goal:"data"`, single metric | Card Row → `goal:"features"`, 3 cards | Comparison Split → `goal:"content"`, two-column contrast | Timeline Horizontal → `component_category:"timeline"` | Quote Spotlight → `goal:"content"`, quote in bullets | Code Terminal → `goal:"code"` | Full-Width Visual → `goal:"content"`, full-bleed image

**Data**: Table Diagram → `goal:"data"`, diagram type:"table" | Chart Focus → `goal:"data"`, diagram type: chart | Metric Dashboard → `component_category:"infographic"` | Infographic Component → `component_type:"group"` | Number Grid → `goal:"data"`, 2×2 metrics

**Content relationship → visual strategy**: Sequential → Timeline | Contrast → Comparison Split | Primary+secondary → unequal layout | Equal-weight → Card Row | Hierarchical → Hierarchy component | Evidence → center + orbit | Process → Cycle/Process component | Data-driven → Big Number/Chart

### Dial Action Map (V/M/D → LLM decisions in content.json)

| VARIANCE | Action |
|----------|--------|
| 1-3 | Centered layouts; equal-width cards; `layout_variant:"centered"` or `"standard"` |
| 4-7 | Left-align titles; feature first card; `layout_variant:"sidebar-left"` or `"asymmetric"`; vary image position |
| 8-10 | Unequal grids; asymmetric hero; vary `layout_variant` per slide; avoid repeated layout families |

| MOTION | Action |
|--------|--------|
| 1-3 | No `animation` entries; slide transitions only |
| 4-7 | Add `animation:"fade-in"` on cover + section dividers; stagger bullets |
| 8-10 | Add `animation:"fly-in"` on most slides; sequence reveals; `animation:"grow"` on CTA |

| DENSITY | Action |
|---------|--------|
| 1-3 | 2-3 bullets/page; breathing slides after every content slide |
| 4-7 | 3-5 bullets/page; mix densities across pages |
| 8-10 | 6+ bullets on data slides; compact tables; infographic components |

### Redesign Protocol
- Greenfield: start from Dial baseline
- Redesign-Preserve: audit brand tokens → incremental evolution
- Redesign-Overhaul: visually equivalent to greenfield
- Audit before modifying: brand tokens / information architecture / content blocks / patterns to keep / patterns to kill
- Modernization levers (in order): fonts → spacing → colors → animations → key-page rebuild → full replacement
- Never silently change: page order / navigation labels / logo / legal copy

### Design System Mapping
- Consulting/finance → sidebar + component_library process/hierarchy
- Tech talks → code block + component_library infographic
- Education → exercise page + built-in bullets
- Creative proposals → custom blocks + AI images
- Brand launch → full-bleed images + minimal text
- ONE design system per deck — no mixing McKinsey sidebar with Cyberpunk neon

### Performance & Accessibility
- <50 shapes per slide | images: cover-fit crop, never stretch | cache-first
- Public-sector / accessibility scenes: motion ≤3 | unknown audience: motion ≤5
- Dark mode: no pure black (#000) or pure white (#FFF) — use near-black/near-white
- Z-order: background < content < decoration < overlay

### Scope Exclusions
- Pure data tables → Excel | Multi-step forms → Web app | Real-time collaboration → Dedicated app | Interactive dashboards → Power BI/Tableau | Long documents (>50 pages) → Word/PDF

### Pre-Flight Check
- [ ] Basics: fonts ≥11pt | pages ≥3 shapes | text on every page | images correct | no broken links
- [ ] Consistency: accent color throughout | corner radius uniform | font pair uniform | theme locked | font-size levels ≥4
- [ ] Typography: cover title ≤2 lines | subtitle ≤20 chars | inner title ≤2 lines | no rotated text | line-height ≥1.1 for italic
- [ ] Layout: adjacent pages different layout family | 8+ pages ≥4 layouts | first card featured | no 3 consecutive same-structure pages | alternation ≤2
- [ ] Labels: eyebrow ≤ceil(pages/3) | no numbered eyebrows | no image overlay labels | no status dots | no generic step labels
- [ ] Color: no default-blue cover (unless specified) | no default gold+navy | chart colors from palette | dark-theme contrast sufficient | no light-gray on white
- [ ] Rhythm: 6+ pages have divider | hook ≠ cta visually | density varies | high-density followed by low | ≤1 core message per page
- [ ] Content: no AI Tells violations | no fake precision numbers | bullets have logical relation | quotes ≤3 lines with attribution
- [ ] Visuals: cover has real visual | minimalist ≥2-3 pages with images | bento ≥2-3 cells varied | logo wall uses images
- [ ] Scene: projection contrast OK | print doesn't rely on animation | large-screen numbers ≥36pt
- [ ] Animation: each has stated purpose | motion>4 has real animations | marquee ≤1/page
- [ ] Dial: values derived from Design Read | variance>4 has asymmetric layouts | density varies across pages
```

## 行数对比

| 章节 | v1 方案 | v2 修正版 | 变化 |
|------|---------|----------|------|
| Typography | 8 | 9 | +1: "each page" rule moved from Data |
| Color | 7 | 7 | 0: 去重后不变 |
| Layout | 9 | 9 | 0: 规格页/20行表从Data移入 |
| Page Roles | 4 | 4 | 0 |
| Data & Content | 8 | 0 | **删除**: 全部合并入 Content Design Rules |
| Visual Assets | 4 | 4 | 0 |
| Consistency Locks | 5 | 3 | -2: 去重 |
| AI Tells | 28 | 25 | -3: 去重 |
| Design Vocabulary | 8 | 10 | +2: 加 content.json 映射列 |
| Dial Action Map | 3 | 18 | +15: 新增动作映射 |
| Redesign Protocol | 6 | 6 | 0 |
| Design System Mapping | 2 | 2 | 0 |
| Performance | 4 | 4 | 0 |
| Scope | 1 | 1 | 0 |
| Pre-Flight Check | 12 | 12 | 0 |
| **Design Constraints 小计** | **~281** | **~265** | **-16** |
| Content Design Rules | 14 | 20 | +6: 合并4条增量规则 |
| **净变化** | | | **-10** |

## 总结：5 个问题的解决方案

| # | 问题 | 解决方案 | 影响 |
|---|------|----------|------|
| 1 | AI Tells 重复 | 去重 3 条（状态点×1、地域时间×1、滚动提示×1合并） | -3 条 |
| 2 | Content Design Rules 重叠 | 删 Design Constraints 中 6 条重复项，升级 Content Design Rules 为 12 行表格 | -6 条，+6 行 CDR |
| 3 | 中英文混用 | Design Constraints 全英文，中文术语括号标注，Vocabulary 加 content.json 映射 | 0 行变化，质量提升 |
| 4 | Dial 无动作映射 | 新增 Dial Action Map（3 个表，~18 行），定义每个 Dial 范围→LLM 在 content.json 中的具体动作 | +15 行 |
| 5 | Consistency Locks 重复 | 删 3 条重复（accent/warm-cool/theme），加交叉引用 | -3 条 |

**修正后 Design Constraints ~265 行 + Content Design Rules 20 行 = ~285 行设计指导。**

相比 v1 的 281+14=295 行，净减 10 行，但消除了 5 个内部矛盾。

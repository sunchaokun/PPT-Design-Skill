# Architecture — Module Registry

**⚠️ READ BEFORE deleting any file.** Every module listed here has a deletion risk classification. Removing a CORE or ENTRY module will break the project.

Run `python scripts/check_deps.py <module_name>` before deleting any module to see full impact.

## Risk Levels

| Level | Meaning | Action |
|-------|---------|--------|
| **CORE** | Imported by ≥3 modules; deletion causes cascade failure | NEVER delete |
| **ENTRY** | Entry point called by `__init__.py`, CLI, or external code; not detected by static import analysis | NEVER delete |
| **INTERNAL** | Imported by 1-2 modules; deletion requires updating callers | CAUTION — update callers first |
| **ISOLATED** | No internal imports; can be removed independently | SAFE — but verify no runtime/CLI usage |

## Phase 1: Planning

| Module | Risk | Purpose |
|--------|------|---------|
| `planner/story_planner.py` | INTERNAL | Narrative planning, page structure, emotion arc |

## Phase 2: Design Decisions

| Module | Risk | Purpose |
|--------|------|---------|
| `decider/design_decider.py` | CORE | Per-page layout/color/typography from 40K+ combos; imported by content_generator, ppt_renderer, qa_gates |

## Phase 3: Content Generation

| Module | Risk | Purpose |
|--------|------|---------|
| `content/content_generator.py` | INTERNAL | Copy formulas (PAS/FAB/AIDA) + image keywords |
| `enterprise/content_parser.py` | INTERNAL | Parses content.json AND README.md; H1→pages, goal inference |

## Phase 4: Rendering

| Module | Risk | Purpose |
|--------|------|---------|
| `renderer/ppt_renderer.py` | ENTRY | Phase 4 PPT rendering; called by `__init__.py` freestyle path |
| `renderer/theme_composer.py` | INTERNAL | 40,000+ style combos; 35 moods |
| `renderer/image_fetcher.py` | INTERNAL | 4 image engines (Seedream/GPT Image/DALL-E/Wanx) + 1 enhancer (Kimi) |
| `renderer/diagram_engine.py` | CORE | 10 diagram types; imported by precision_renderer, block_renderer, component_renderer |
| `renderer/diagram/base.py` | CORE | Diagram base class; imported by ALL 10 diagram types + diagram_engine |
| `renderer/diagram/layout_engine.py` | CORE | Diagram layout computation; imported by 5 modules |
| `renderer/diagram/diagram_style.py` | CORE | Diagram styling; imported by 5 modules |
| `renderer/diagram/flowchart.py` | INTERNAL | Flowchart diagram |
| `renderer/diagram/funnel.py` | INTERNAL | Funnel diagram |
| `renderer/diagram/timeline.py` | INTERNAL | Timeline diagram |
| `renderer/diagram/swot.py` | INTERNAL | SWOT diagram |
| `renderer/diagram/hierarchy.py` | INTERNAL | Hierarchy diagram |
| `renderer/diagram/matrix.py` | INTERNAL | Matrix diagram |
| `renderer/diagram/cycle.py` | INTERNAL | Cycle diagram |
| `renderer/diagram/pyramid.py` | INTERNAL | Pyramid diagram |
| `renderer/diagram/venn.py` | INTERNAL | Venn diagram |
| `renderer/diagram/table.py` | INTERNAL | Table diagram |
| `renderer/diagram/connector_router.py` | ISOLATED | Connector path routing |
| `renderer/diagram/text_measurer.py` | ISOLATED | Text measurement for diagrams |
| `renderer/diagram/data_splitter.py` | ISOLATED | Data splitting for diagrams |
| `renderer/visual_effects.py` | CORE | Shape3D, bevel, 31 patterns, frosted glass; imported by 7 modules |
| `renderer/text_effects.py` | INTERNAL | Gradient fill, outline, shadow, glow, 3D, vertical text, rotation, alpha, spacing |
| `renderer/blip_fill.py` | INTERNAL | Image-in-shape via blipFill; circle/hexagon/diamond image, 22 artistic effects |
| `renderer/image_processor.py` | ISOLATED | 7 Pillow filters (grayscale, sepia, duotone, ink_wash, blur, vignette, edge_fade) |
| `renderer/animation.py` | INTERNAL | 12 transitions, 11 entrances, 8 exit presets, 8 emphasis presets, morph |
| `renderer/decoration_library.py` | INTERNAL | 7 decorations: brush divider, seal stamp, scroll, neon, grid, glass, ink splash |
| `renderer/chart_builder.py` | INTERNAL | Chart construction utilities |
| `renderer/color_system.py` | ISOLATED | Color system definitions |
| `renderer/decoration_renderer.py` | ISOLATED | Decoration rendering |
| `renderer/effects.py` | ISOLATED | Effect utilities |
| `renderer/elevation.py` | ISOLATED | Elevation/shadow system |
| `renderer/freeform_builder.py` | ISOLATED | Freeform shape builder |
| `renderer/group_builder.py` | ISOLATED | GroupShape builder |
| `renderer/layout_engine.py` | ISOLATED | Layout computation engine |
| `renderer/layout_registry.py` | CORE | Layout type registry; imported by ppt_renderer, precision_renderer, enterprise_renderer |
| `renderer/shape_factory.py` | INTERNAL | Shape creation factory |
| `renderer/shape_utils.py` | ISOLATED | Shape utility functions |
| `renderer/theme_mapper.py` | INTERNAL | Theme mapping (design system → PPT theme) |
| `renderer/typography.py` | ISOLATED | Typography definitions |

## Enterprise Pipeline

| Module | Risk | Purpose |
|--------|------|---------|
| `enterprise/pipeline.py` | ENTRY | Orchestration: always uses PrecisionRenderer; run_beautify() for beautify mode |
| `enterprise/precision_renderer.py` | CORE | Unified renderer; render_slide() dispatches by goal; imported by pipeline, proposal_generator, component_renderer, design_dna_extractor |
| `enterprise/brand_spec.py` | CORE | Brand spec data class; imported by 8 modules (pipeline, precision_renderer, component_adapter, etc.) |
| `enterprise/brand_color_context.py` | ISOLATED | Brand color context for component adaptation |
| `enterprise/scanner.py` | ISOLATED | Scans project dir for template.pptx, brand.json, content.json, README.md, images |
| `enterprise/slide_utils.py` | CORE | Slide utility functions; imported by 4 modules |
| `enterprise/content_parser.py` | INTERNAL | (listed in Phase 3; dual-category) |
| `enterprise/image_matcher.py` | ISOLATED | Image assignment: keyword match + size classification + auto_generate_image_prompts() |
| `enterprise/proposal_generator.py` | ENTRY | Generate 2-3 style preview PPTs; called by __init__.py proposal flow |
| `enterprise/slide_extractor.py` | ISOLATED | Extract content + layout from existing PPT (beautify mode) |
| `enterprise/smartart_extractor.py` | ISOLATED | SmartArt XML parsing (4 parts: data/layout/colors/quickStyle) |
| `enterprise/group_extractor.py` | ISOLATED | GroupShape recursive extraction |
| `enterprise/ole_extractor.py` | ISOLATED | OLE/embedded object metadata extraction |
| `enterprise/component_library.py` | CORE | SQLite-indexed component library; imported by pipeline, component_renderer, design_dna_extractor |
| `enterprise/component_renderer.py` | CORE | Component rendering bridge; imported by precision_renderer, component_library, design_dna_extractor |
| `enterprise/component_adapter.py` | INTERNAL | Component color adaptation (fill_map/text_map) |
| `enterprise/block_renderer.py` | INTERNAL | Block-level rendering (badge, table, etc.) |
| `enterprise/enterprise_renderer.py` | ISOLATED | Legacy enterprise renderer (deprecated, kept for compatibility) |
| `enterprise/enterprise_decider.py` | ISOLATED | Enterprise-specific design decisions |
| `enterprise/design_dna_extractor.py` | INTERNAL | Design DNA extraction from existing PPT |
| `enterprise/template_analyzer.py` | ISOLATED | Template analysis for VI Build |
| `enterprise/delivery_gate.py` | ISOLATED | Delivery quality gate checks |
| `enterprise/density_profile.py` | ISOLATED | Density profile configuration |
| `enterprise/review_gate.py` | ISOLATED | Review gate checks |
| `enterprise/page_revision.py` | ISOLATED | Page revision (add/delete/swap/move) |
| `enterprise/version_manager.py` | ISOLATED | Version control for output files |

## Build Helpers

| Module | Risk | Purpose |
|--------|------|---------|
| `build_helpers.py` | ENTRY | LLM build script toolbox; 39 functions, Typography/Spacing classes; called via `from ppt_pro_max.build_helpers import *` |

## Adapters

| Module | Risk | Purpose |
|--------|------|---------|
| `adapters/ui_ux_adapter.py` | CORE | UI/UX Pro Max adapter; imported by 4 modules (theme_composer, slide_search_adapter, story_planner, design_decider) |
| `adapters/llm_config_adapter.py` | ISOLATED | LLM configuration adapter |
| `adapters/slide_search_adapter.py` | ISOLATED | Slide search adapter |

## CLI & Scripts

| Module | Risk | Purpose |
|--------|------|---------|
| `cli.py` | ENTRY | Command-line interface entry point |
| `analyze_template.py` | ENTRY | Template analysis CLI tool (`python -m ppt_pro_max.analyze_template`) |
| `scripts/build_library.py` | ISOLATED | Component library build script |
| `qa/qa_gates.py` | INTERNAL | Quality assurance gate checks |

## Dependency Hotspots (CORE modules — NEVER delete)

```
brand_spec ──────── 8 importers (pipeline, precision_renderer, component_adapter, component_renderer, proposal_generator, enterprise_decider, template_analyzer, design_dna_extractor)
diagram/base ────── 11 importers (ALL 10 diagram types + diagram_engine)
visual_effects ──── 7 importers (build_helpers, precision_renderer, diagram/base, shape_factory, color_system, elevation, decoration_library)
diagram_style ───── 5 importers (block_renderer, precision_renderer, diagram/base, diagram_engine, component_renderer)
layout_engine ───── 5 importers (block_renderer, precision_renderer, diagram/base, diagram_engine, component_renderer)
diagram_engine ──── 3 importers (block_renderer, precision_renderer, component_renderer)
precision_renderer ─ 4 importers (pipeline, proposal_generator, component_renderer, design_dna_extractor)
slide_utils ─────── 4 importers (page_revision, design_dna_extractor, precision_renderer, enterprise_renderer)
layout_registry ─── 3 importers (ppt_renderer, precision_renderer, enterprise_renderer)
ui_ux_adapter ───── 4 importers (theme_composer, slide_search_adapter, story_planner, design_decider)
component_library ─ 3 importers (pipeline, component_renderer, design_dna_extractor)
component_renderer ─ 3 importers (precision_renderer, component_library, design_dna_extractor)
design_decider ──── 3 importers (content_generator, ppt_renderer, qa_gates)
```

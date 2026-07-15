# PPT-Design-Skill 修订方案

## 目标

用户只需说"做个国际风的PPT"+ 丢素材，系统自动：方案预览 → 用户确认 → 完整生成。

---

## 一、当前问题

| # | 问题 | 影响 | 代码现状 |
|---|------|------|----------|
| 1 | Pipeline 有模板时部分内容用 paragraph 级字体（`p.font.size`/`p.font.color.rgb`），PowerPoint 不认 paragraph 级字体设置 | 字体全回退为宋体 | `pipeline.py:1010-1021`（cards）、`pipeline.py:1148-1162`（code block）、`pipeline.py:1208-1227`（exercise）用 `p.font`；`enterprise_renderer.py:181-183`（页码）用 `p.font`。`_populate_slide()` 的 title/bullets 已用 run-level（`run.font`），但 cards/code/exercise 未改 |
| 2 | 内容图片保持宽高比但不裁剪，留白/不填满；Hero 图片路径已有 Pillow 裁剪 | 内容图留白多 | `pipeline.py:898-919` 的 `_insert_content_image()` 用 Pillow 读尺寸计算宽高比后 `add_picture`，不裁剪；`pipeline.py:795-834` 的 Hero 路径有 Pillow cover-crop |
| 3 | 布局由母版模板决定，无法按内容自适应 | 内容多时溢出，内容少时空洞 | `enterprise_decider.py:26-29` 仅做 `layout_mapping` 查找，无内容量自适应逻辑；`density_profile.py` 调节字号/max_bullets/max_bullet_chars/line_spacing/image_width_ratio，但不改变布局结构 |
| 4 | 有模板时 Pipeline 的 cards/code/exercise 用 paragraph 级字体，无模板才走 PrecisionRenderer（run-level 字体） | 有模板的企业场景 cards/code/exercise 字体回退 | `pipeline.py:139-164`：`use_freestyle_render = not renderer.has_template`，有模板走 `_populate_slide()`（title/bullets 已用 run-level，但 cards/code/exercise 仍用 paragraph 级），无模板走 `_render_with_precision_renderer()`（run 级） |
| 5 | AI 配图默认 `image_mode=placeholder`，ImageFetcher 不构建；需手动配 `--llm-provider`/`--llm-api-key` 且 image_prompt 需手写 | 默认无配图 | `pipeline.py:1362-1363`：`image_mode` 默认 `"placeholder"` 时 `ImageFetcher` 返回 `None`；`image_fetcher.py:165-197` 的 `_build_image_prompt()` 已能自动生成 prompt，但默认模式下不触发 |
| 6 | 无方案预览机制，一次出完整 PPT | 用户无法提前确认风格方向 | `review_gate.py` 仅生成 JSON proposal（文字描述），不生成 PPT 预览文件 |
| 7 | 内容结构需手写 content.json | 用户不会写 | `content_parser.py` 仅解析 content.json，无 README.md 解析能力；`scanner.py` 也不扫描 README.md |

---

## 二、修订架构

```
用户输入
  │ (主题描述 + 素材目录)
  ↓
┌─────────────────────────────┐
│ ① 内容智能解析                │  README/大纲/素材 → pages[] 结构
│    ContentParser             │  自动推断页数、goal、bullets/cards/diagram
│    (扩展现有 content_parser)  │
└──────────┬──────────────────┘
           ↓
┌─────────────────────────────┐
│ ② 方案生成                   │  ThemeComposer.compose(style) × 2~3 组
│    ProposalGenerator         │  每组生成 4 页预览（封面/文字内页/数据内页/CTA）
│    (新增模块)                 │  保存为 proposal_A.pptx / B / C
└──────────┬──────────────────┘
           ↓
      用户确认方案
           ↓
┌─────────────────────────────┐
│ ③ 完整生成                   │  用确认的原子 + 完整 pages[]
│    PrecisionRenderer         │  render_slide() 自动模板分发
│    (扩展 render_slide)       │  图片：本地优先 → AI 补缺
│                              │  品牌合规：LOGO/页脚/水印
└──────────┬──────────────────┘
           ↓
        output.pptx
```

---

## 三、各模块详细设计

### 3.1 PrecisionRenderer.render_slide()

**输入**：一个 page dict（同现有 Pipeline 的 page_designs 结构）

```python
page = {
    "goal": "hook",           # hook/cta/problem/solution/features/market/data/...
    "title": "Zensers",
    "subtitle": "AI 驱动的研究平台",
    "bullets": ["要点1", "要点2"],
    "image": "/path/to/img.jpg",     # 本地图片（优先）
    "image_prompt": "...",           # AI 配图描述（缺口时自动生成）
    "cards": [{"title": "...", "text": "..."}],
    "diagram_type": "funnel",
    "diagram_data": {...},
    "code": {"language": "python", "source": "..."},
    "exercise": {...},
    "chart": {...},                  # 图表数据
    "notes": "...",                  # 演讲者备注
    "links": [...],                  # 超链接
}
```

> **注意**：此结构已与 Pipeline 的 `_build_page_designs()` 输出一致（`pipeline.py:354-377`），无需新增字段，仅 `image_prompt` 为新增。

**自动模板分发逻辑**（基于现有 `layout_registry.py` 的 goal_mapping）：

```
goal=hook/cta → title-slide / cta-closing 模板（全幅图+遮罩+大标题+副标题）
goal=problem/solution/features/content/market/business/team/financial/agitation/proof → content-with-title 模板（10 个 goal 共享，有 image 则右侧配图）
goal=solution → 也匹配 three-column-cards（3 cards）
goal=features → 也匹配 three-column-cards（3 cards）或 grid-2x2-cards（4 cards）
goal=data/financial → chart-focus 模板
goal=code → code-block 模板（goal_mapping 为 ["code", "technical", "demo"]）
goal=exercise → exercise-layout 模板
goal=navigation/agenda/overview → sidebar-left 模板
goal=funnel/pipeline/conversion → funnel 模板
goal=content → 根据 bullets/cards/image 自动选：
               有 cards → three-column-cards / grid-2x2-cards
               有 image + bullets → image-plus-text 模板
               只有 bullets → content-with-title 模板
```

> **注意**：模板名称需与 `layout_registry.py` 的 `MASTER_LAYOUTS` key 一致，文档原版使用的 `numbered_list`/`process_flow`/`3cards`/`4cards`/`table_of_contents`/`bullet_list` 在代码中不存在。

**关键原则**：
- 所有文本 run-level 字体（PrecisionRenderer 已实现：`add_text()`/`add_multiline()` 均用 `run.font`）
- 所有图片 Pillow 预裁剪（PrecisionRenderer 已实现：`add_image()` 有 Pillow crop）
- 品牌约束自动应用：accent 竖条、muted 底条、LOGO、页脚、水印（PrecisionRenderer 已实现：`apply_brand_background()`/`apply_logo()`/`apply_footer()`/`apply_watermark()`）
- 内容自适应：bullets 少→大字少行，bullets 多→小字多行（需新增，当前 PrecisionRenderer 无此逻辑）
- goal 是默认模板映射，page dict 可显式指定 `layout` 覆盖
- 有模板时：仅借用母版背景和配色方案，内容定位与无模板一致（不填 placeholder），见 3.5 节

---

### 3.2 内容智能解析 ContentParser

**输入**：
- `query`: 主题描述（"Zensers 产品介绍"）
- `materials_dir`: 素材目录路径
- `business_mode`: 可选（pitch/education/training/report）

**输出**：`pages[]` 列表（同现有 `content_parser.py` 的输出格式）

**逻辑**：
1. 如果有 content.json → 直接用现有 `load_enterprise_content()`（最高优先级，已实现）
2. 如果有 README.md → 解析为页面结构（**新增**）：
   - 一级标题 → 新页面
   - 二级标题 → page 的 bullets/cards
   - 表格 → diagram_type=table
   - 代码块 → code
   - 图片引用 → 分配到对应页面
3. 如果只有 query → 调用现有 `StoryPlanner.plan()` 生成大纲，再补充内容（Pipeline 已实现此路径：`pipeline.py:299-323`）

**页面结构推断**：
- 第一页 → goal=hook（封面）
- 第二页 → goal=overview（sidebar-left 模板，目录/议程页）
- 后续根据内容语义推断 goal：
  - "痛点/问题/挑战" → problem
  - "方案/解决/如何" → solution
  - "功能/特性/能力" → features
  - "市场/规模/TAM" → market（映射到 content-with-title，非 funnel）
  - "架构/技术/系统" → data
  - "代码/示例/快速开始" → code
  - "开始/联系/下一步" → cta

> **注意**：现有 `StoryPlanner` 的策略结构（`story_planner.py:28-65`）中无 `toc` goal，Education Course 用 `overview`，Business Report 也用 `overview`。`sidebar-left` 的 goal_mapping 为 `["navigation", "agenda", "overview"]`，不含 `toc`。需在 StoryPlanner 的策略中增加 `toc` goal 或在 ContentParser 中将 `toc` 映射为 `overview`。

---

### 3.3 图片优先逻辑

**规则**：
1. 扫描项目目录下的所有图片文件（png/jpg/jpeg/webp/svg）——现有 `scanner.py:12` 已定义 `_IMAGE_EXTENSIONS`，但仅扫描项目根目录，不递归子目录
2. Logo → 单独识别（文件名含 logo/brand）——现有 `scanner.py:14` 已实现 `_LOGO_PATTERN`
3. 其余图片按尺寸分类（**新增**，现有 `image_matcher.py` 仅按文件名关键词匹配，不按尺寸分类）：
   - 大图（>1500px 宽） → 背景候选（分配到 hook/cta 页）
   - 中图（800-1500px） → 产品/场景图（分配到 features/solution 页）
   - 小图（<800px） → 图标/缩略图（分配到 cards 内）
4. 分类后按页面缺口均匀分配
5. 未分配到图的页面 → 自动生成 image_prompt（**新增**，现有 `ImageFetcher._build_image_prompt()` 已能根据 goal+emotion 生成 prompt，但 Pipeline 未在缺图时自动调用）：
   - goal=hook → "professional presentation cover, {title}, minimalist, high quality"
   - goal=problem → "business challenge visualization, dark tones, cinematic"
   - goal=solution → "technology solution dashboard, clean, modern"
   - goal=features → "product features grid, bright, professional"
   - goal=market → "market growth chart, data visualization, blue tones"
   - goal=cta → "professional contact slide, clean, inviting"
6. image_prompt 传给 ImageFetcher → AI 生成补充

> **注意**：文档原版有两个编号"5"，已修正为 5 和 6。

---

### 3.4 方案预览 ProposalGenerator

**流程**：
1. 从 ThemeComposer 获取 2~3 组设计原子（基于用户风格描述）
2. 每组生成 4 页预览 PPT：
   - 封面页（goal=hook）
   - 文字内页（goal=problem，展示 bullets+image 布局）
   - 数据内页（goal=features，展示 cards 布局）
   - CTA 页（goal=cta，展示结尾页风格）
   注：4 页比 3 页更好——用户能看到文字页、数据页和结尾页三种风格的呈现效果
3. 保存为：
   ```
   output/
   ├── proposal_A.pptx   (方案A: 如 slate-minimal + modern-sans + no-decoration)
   ├── proposal_B.pptx   (方案B: 如 ocean-blue + geometric-sans + accent-bar)
   └── proposal_C.pptx   (方案C: 如 midnight-navy + clean-corporate + sidebar-nav)
   ```
4. 返回方案描述供用户选择

**方案差异化策略**：
- 方案A：用户描述最接近的风格
- 方案B：同 mood 但不同 palette（如冷色→暖色）
- 方案C：不同 mood 的替代方案（如 professional → creative）

> **注意**：文档原版写"3 页预览（封面/目录/内页）"与后面"4 页"矛盾，已统一为 4 页。目录页（goal=toc）在预览中意义不大（无实际内容），替换为 CTA 页更实用。

---

### 3.5 Pipeline 修订

**核心改动**：Pipeline **始终**走 PrecisionRenderer

```
当前 (pipeline.py:136-249):
  has_template → EnterpriseRenderer._populate_slide() (paragraph 级字体，低质量)
  no_template  → PrecisionRenderer._render_with_precision_renderer() (run 级字体，高质量)

修改后:
  has_template → PrecisionRenderer.render_slide() (run 级字体 + 借用模板母版风格)
  no_template  → PrecisionRenderer.render_slide() (run 级字体 + 自由定位)
```

> **注意**：当前 Pipeline 实际有四条渲染路径（`pipeline.py:380-772`）：
> 1. `_populate_slide()` — 有模板时走（title/bullets 用 run 级，cards/code/exercise 用 paragraph 级）
> 2. `_render_with_precision_renderer()` — 无模板时走（run 级，高质量）
> 3. `_render_with_ppt_renderer()` — 已定义但未被调用（死代码）
> 4. `_populate_slide_with_layout()` — 已定义但未被调用（用 LayoutRegistry，预留路径）
> 
> 修订后应删除路径 1/3/4，统一走 PrecisionRenderer。

有模板时 PrecisionRenderer 的行为——**只借风格，不填 placeholder**：

用户给模板是让我们参考设计风格（母版背景、配色、字体、尺寸），不是让内容去迁就模板的 placeholder 框。死板填充 placeholder 会导致内容被模板结构束缚，丢失自适应能力。

- `create_presentation()` 从模板创建——借用母版、配色方案、幻灯片尺寸（已实现 `precision_renderer.py:71-84`）
- `add_slide()` 用模板的 blank layout 或最简 layout——仅借用母版背景装饰，不绑定 placeholder 框架（已实现 `precision_renderer.py:86-105`，当前逻辑优先选 title layout，应改为优先选 blank layout）
- 内容定位：与无模板时一致，由 PrecisionRenderer 按 goal→layout 分发 + 精确坐标自由定位（已实现）
- 品牌色覆盖：从模板提取的 BrandSpec 颜色优先，确保文字/形状与模板风格协调（已实现）
- 图片仍用 Pillow 预裁剪（已实现）
- LOGO/页脚/水印仍由 BrandSpec 控制（已实现）

**本质**：有模板和无模板的唯一区别是 `create_presentation()` 的来源——从模板创建则自带母版背景和配色方案，从空白创建则靠 ThemeComposer 提供。内容渲染逻辑完全一致，不做 placeholder 填充。

---

## 四、新增/修改文件清单

| 文件 | 操作 | 说明 |
|------|------|------|
| `src/ppt_pro_max/enterprise/precision_renderer.py` | **修改** | 新增 `render_slide()` 方法，整合 goal→layout 分发；`add_slide()` 有模板时优先选 blank layout |
| `src/ppt_pro_max/enterprise/content_parser.py` | **修改** | 扩展：新增 README.md 解析能力（原仅有 content.json 解析） |
| `src/ppt_pro_max/enterprise/image_matcher.py` | **修改** | 扩展：新增按尺寸分类逻辑（原仅按文件名关键词匹配） |
| `src/ppt_pro_max/enterprise/proposal_generator.py` | **新增** | 方案预览生成（2-3 个方案 × 4 页） |
| `src/ppt_pro_max/enterprise/pipeline.py` | **修改** | 统一走 PrecisionRenderer + 方案流程；删除 `_populate_slide()`、`_render_with_ppt_renderer()` 和 `_populate_slide_with_layout()` 三条旧路径 |
| `src/ppt_pro_max/renderer/theme_composer.py` | **修改** | 扩展 mood_words（international/cream/frosted/mckinsey 等） |
| `src/ppt_pro_max/__init__.py` | **修改** | generate_ppt() 新增 proposal/confirmed_proposal/materials_dir 参数 |

> **注意**：文档原版将 `content_parser.py` 标为"新增 content_intelligence.py"，但现有 `content_parser.py` 已有基础解析能力，应扩展而非新建。同理 `image_matcher.py` 已存在，应扩展而非新建 `image_planner.py`。原版 `proposal_generator.py` 的说明写"2-3 个方案×3 页"，与 3.4 节的 4 页矛盾，已统一为 4 页。

---

## 五、generate_ppt() API 变更

```python
# 现有：直接生成
result = generate_ppt("Zensers 产品介绍", style="international")

# 新增：方案预览模式
result = generate_ppt(
    "Zensers 产品介绍",
    style="international",
    proposal=True,              # 生成 2-3 个预览方案
    materials_dir="./assets",   # 素材目录
)
# 返回: {
#   "proposals": [
#     {"id": "A", "path": "output/proposal_A.pptx", "atoms": {...}},
#     {"id": "B", "path": "output/proposal_B.pptx", "atoms": {...}},
#     {"id": "C", "path": "output/proposal_C.pptx", "atoms": {...}},
#   ]
# }

# 新增：确认方案后完整生成
result = generate_ppt(
    "Zensers 产品介绍",
    style="international",
    confirmed_proposal="B",     # 用户选了方案B
    materials_dir="./assets",
)
# 返回: {"output_path": "...", "num_slides": 14, "render_mode": "precision"}
```

> **注意**：现有 `generate_ppt()` 的 `project` 参数触发 enterprise pipeline，`materials_dir` 是新增参数，需与 `project` 区分——`project` 指向完整项目目录（含 template.pptx/brand.json 等），`materials_dir` 仅指向素材目录（图片/README 等）。两者可同时使用：`project` 提供模板和品牌规范，`materials_dir` 提供内容素材。

---

## 六、实施顺序

| 阶段 | 内容 | 依赖 | 涉及文件 | 状态 |
|------|------|------|----------|------|
| **P1** | PrecisionRenderer.render_slide() | 无 | `precision_renderer.py` | **已完成** |
| **P2** | Pipeline 统一走 PrecisionRenderer | P1 | `pipeline.py` | **已完成** |
| **P3** | mood_words 扩展 | 无 | `theme_composer.py` | **已完成** |
| **P4** | 内容智能解析 ContentParser（README 解析） | 无 | `content_parser.py`, `scanner.py` | **已完成** |
| **P5** | 图片尺寸分类 + image_prompt 自动生成 | P4 | `image_matcher.py` | **已完成** |
| **P6** | 方案预览 ProposalGenerator | P1+P3+P4+P5 | `proposal_generator.py`（新增） | **已完成** |
| **P7** | generate_ppt() API 集成 | P6 | `__init__.py` | **已完成** |
| **P8** | 端到端测试 | P7 | `tests/` | **已完成** |

P1-P3 是基础，先做；P4-P5 让系统"能理解用户"；P6-P7 让系统"能和用户互动"；P9 让系统"能改别人的活"。

---

## 八、美化改版（Beautify 模式）

### 场景

客户给一份已有 PPT（内容完整但设计粗糙），要求美化。核心区别：

| | 从零生成 | 美化改版 |
|---|---------|---------|
| 内容来源 | StoryPlanner 生成 / content.json 提供 | 从客户 PPT 的每一页提取 |
| 内容修改 | 系统决定 | **不改动**，只改视觉呈现 |
| 布局 | PrecisionRenderer 自由定位 | **保留原布局结构**，仅升级视觉 |
| 风格 | ThemeComposer 选择 | ThemeComposer 选择，但需兼容原内容量 |

### 流程

```
客户 PPT (ugly.pptx)
  │
  ↓
┌─────────────────────────────┐
│ ① 内容提取                    │  逐页提取：标题、正文、图片、图表
│    SlideExtractor             │  输出 pages[]（同 render_slide 输入格式）
│                               │  同时提取每页的布局结构信息
└──────────┬──────────────────┘
           ↓
┌─────────────────────────────┐
│ ② 风格选择                   │  同方案预览流程
│    ProposalGenerator         │  2~3 组风格方案 × 4 页预览
└──────────┬──────────────────┘
           ↓
      用户确认风格
           ↓
┌─────────────────────────────┐
│ ③ 美化渲染                   │  逐页：原内容 + 新风格 → render_slide()
│    PrecisionRenderer         │  保留原布局结构（内容位置/比例），
│    (beautify=True)           │  升级：字体、配色、装饰、图片质量
└──────────┬──────────────────┘
           ↓
      beautified.pptx
```

### 8.1 SlideExtractor — 从已有 PPT 提取内容

**输入**：客户 PPT 文件路径

**输出**：`pages[]` + 每页的 `layout_hint`

```python
page = {
    "goal": "content",              # 从内容语义推断
    "title": "市场分析",             # 提取的标题
    "subtitle": None,
    "bullets": ["市场规模达千亿", "年增长率 25%"],
    "image": None,                   # 提取的图片（保存到临时目录）
    "cards": [],
    "diagram_type": None,
    "diagram_data": None,
    "code": None,
    "exercise": None,
    "chart": None,
    "notes": "强调增长趋势",          # 提取的演讲者备注
    "links": [],
    # ── 美化模式专有字段 ──
    "layout_hint": {                 # 原始布局结构，供渲染时参考
        "title_pos": (0.9, 0.5),     # 标题原始位置（英寸）
        "body_pos": (0.9, 1.6),      # 正文原始位置
        "image_pos": (8.5, 1.6),     # 图片原始位置
        "has_image": True,
        "bullet_count": 2,
        "is_full_bleed": False,      # 是否全幅图页
    },
    "complex_elements": [            # 无法标准化的复杂元素
        {
            "type": "smartart",      # smartart / group / ole / chart_embed
            "category": "process",   # 从 SmartArt 类型推断：process/hierarchy/cycle/matrix/pyramid/relationship/picture
            "texts": ["需求分析", "方案设计", "开发测试", "上线运维"],
            "xml_parts": {...},       # 4 个 XML 快照（data/layout/colors/quickStyle），用于组件库匹配和注入
            "bounds": (0.9, 1.5, 11.5, 4.5),  # 位置和尺寸
        },
    ],
}
```

**提取逻辑**：

1. 遍历每页 slide 的 shapes
2. 识别 shape 类型：
   - `slide.shapes.title` → title
   - placeholder type=BODY/OBJECT → bullets（按 paragraph 拆分）
   - placeholder type=SUBTITLE → subtitle
   - `shape.shape_type == PICTURE` → image（保存到临时目录，记录路径）
   - `shape.shape_type == CHART` → chart（提取图表数据）
   - `shape.shape_type == TABLE` → diagram_type=table
   - `shape.shape_type == GROUP` → 递归提取内部文本和图片 → complex_elements
   - SmartArt → 通过 XML 解析提取文本和类型 → complex_elements（见 8.7）
   - OLE/嵌入对象 → 通过 XML 解析提取元信息 → complex_elements（见 8.7）
3. 记录每个 shape 的位置和尺寸 → `layout_hint`
4. 推断 goal：
   - 第一页 + 有大图/无正文 → hook
   - 最后一页 + 标题含"谢谢/联系/Thank" → cta
   - 标题/正文含关键词 → 按语义推断（同 3.2 的推断规则）
   - 默认 → content
5. 提取 `slide.notes_slide` → notes

**图片处理**：
- 从 PPT 中提取的图片通过 `shape.image.blob` 保存到临时目录
- 记录原始尺寸，用于 `layout_hint` 中的位置参考

### 8.2 美化渲染策略

美化模式的核心原则：**内容不变，视觉升级**。

**逐页渲染逻辑**：

```
beautify 模式下 render_slide() 的行为：

1. 读取 page["layout_hint"]，判断原始布局类型：
   - is_full_bleed=True → hero 布局（全幅图+遮罩+标题）
   - has_image + bullets → image_plus_text 布局
   - 只有 bullets → content-with-title 布局
   - 有 chart → chart-focus 布局
   - ...（同 3.1 的 goal→layout 分发，但优先参考 layout_hint）

2. 内容定位：
   - 保留原始内容的大致位置关系（标题在上、正文在中、图片在右等）
   - 但用 PrecisionRenderer 的精确坐标重新对齐（消除原 PPT 的偏移/不对齐）
   - 字号根据 density_profile 调整（原 PPT 字号混乱时统一）

3. 视觉升级（与从零生成一致）：
   - 字体 → BrandSpec / ThemeComposer 指定（run-level）
   - 配色 → BrandSpec / ThemeComposer 指定
   - 装饰 → accent 竖条、muted 底条、标题下划线
   - 图片 → Pillow 预裁剪（消除拉伸）
   - 背景 → 纯色/渐变（替换原 PPT 的杂乱背景）
   - LOGO/页脚/水印 → 按 BrandSpec 添加

4. 不做的事：
   - 不删减内容（保留客户原始内容（文字、图片、图表）
   - 不改变页码顺序
   - 不新增页面
   - 不替换客户的图片（除非客户明确要求）
```

### 8.3 与从零生成的差异

| 环节 | 从零生成 | 美化改版 |
|------|---------|---------|
| 内容来源 | StoryPlanner / content.json / README | SlideExtractor 从 PPT 提取 |
| goal 推断 | StoryPlanner 策略决定 | 从标题/正文语义推断 |
| 布局选择 | goal → layout_registry | layout_hint 优先，goal 兜底 |
| 内容量 | 系统控制（density 调节） | **固定**（客户原内容，不增删） |
| 图片 | 本地素材 / AI 生成 | **保留原图**，仅裁剪修正 |
| 方案预览 | 同 | 同 |
| 页面增删 | 支持 | 不支持（保持原页数） |

### 8.4 API 设计

```python
# 美化改版模式
result = generate_ppt(
    query="美化这份PPT",
    beautify="client-presentation.pptx",   # 客户 PPT 路径
    style="professional",
)
# 返回: {"output_path": "...", "num_slides": 14, "mode": "beautify"}

# 美化 + 方案预览
result = generate_ppt(
    query="美化这份PPT",
    beautify="client-presentation.pptx",
    style="professional",
    proposal=True,
)
# 返回: {"proposals": [...]}

# 美化 + 确认方案
result = generate_ppt(
    query="美化这份PPT",
    beautify="client-presentation.pptx",
    style="professional",
    confirmed_proposal="A",
)
```

### 8.5 新增文件

| 文件 | 操作 | 说明 |
|------|------|------|
| `src/ppt_pro_max/enterprise/slide_extractor.py` | **新增** | 从已有 PPT 逐页提取内容 + 布局信息 |
| `src/ppt_pro_max/enterprise/precision_renderer.py` | **修改** | `render_slide()` 支持 `layout_hint` 参数 |
| `src/ppt_pro_max/enterprise/pipeline.py` | **修改** | 新增 beautify 分支 |
| `src/ppt_pro_max/__init__.py` | **修改** | `generate_ppt()` 新增 `beautify` 参数 |
| `src/ppt_pro_max/cli.py` | **修改** | 新增 `--beautify` 参数 |

### 8.6 实施顺序

| 阶段 | 内容 | 依赖 |
|------|------|------|
| **P9** | SlideExtractor + beautify 渲染 | P1（需 render_slide） |
| **P10** | beautify API + CLI 集成 | P9 |
| **P11** | 美化模式端到端测试 | P10 |
| **P12** | SmartArt/GroupShape XML 提取器 | P9 |
| **P13** | 组件库基础设施（Schema + Index + CRUD） | P12 |
| **P14** | 组件库匹配引擎 + 渲染桥接 | P13 |

### 8.7 复杂元素提取与组件库系统

#### 验证结论（已通过实测）

**测试 1：合成 SmartArt/GroupShape（test_smartart_inject.py / test_groupshape_extract.py）**

| 元素类型 | python-pptx 高层 API | lxml 低层 XML 解析 | 验证结果 |
|---------|---------------------|-------------------|---------|
| SmartArt 文本 | **不可见**（0 shapes） | `ppt/diagrams/dataN.xml` → `dgm:pt/a:t` | ✓ 4/4 文本提取成功 |
| SmartArt 类型 | **不可见** | `ppt/diagrams/layoutN.xml` → `uniqueId` + `dgm:alg/@type` | ✓ 识别出 lin（线性流程） |
| GroupShape 文本 | `shape.shapes` 可遍历 | `shape._element.findall(".//a:t")` | ✓ 3/3 文本提取成功 |
| GroupShape 结构 | `shape.shapes` 返回子列表 | XML 层级遍历 grpSp/sp | ✓ 递归提取成功 |

**测试 2：真实 SmartArt PPT（art.pptx，3 个 SmartArt 图形）**

| Slide | SmartArt 类型 | 布局 uniqueId | 节点数 | 文本提取 | 类型识别 |
|-------|-------------|--------------|--------|---------|---------|
| 1 | List（蛇形列表） | `.../layout/default` | 5 | ✓ 5/5 | ✓ alg=snake |
| 2 | Relationship（关系图） | `.../layout/target3` | 9 | ✓ 9/9 | ✓ category=relationship |
| 3 | Process（阶段流程） | `.../layout/PhasedProcess` | 9 | ✓ 9/9 | ✓ category=process |

**测试 3：端到端 Round-Trip（test_smartart_roundtrip.py）**

从 art.pptx 提取 SmartArt → 修改配色+文本 → 注入新 PPT → PowerPoint 打开验证。

| 操作 | 修改目标 | 关键发现 | 验证结果 |
|------|---------|---------|---------|
| 文本替换 | `dataN.xml` 的 `dgm:pt/a:t` | pt 的 type 属性为空（非 "node"），需按"有 a:t 且非 doc/parTrans/sibTrans/pres"筛选 | ✓ data 文本替换成功 |
| 配色替换 | `colorsN.xml` 的 `schemeClr` | **PowerPoint 渲染 SmartArt 时从 colors XML 读取 schemeClr 再映射主题色**。必须将 colors XML 的 schemeClr 替换为 srgbClr | ✓ 131 个 schemeClr 全部替换为 srgbClr |
| 注入新 PPT | slide XML + rels + Content_Types | diagram MIME 类型必须精确匹配；slide rels 需 5 个 rId（rId1=slideLayout + rId2-5=dm/lo/qs/cs）；graphicFrame 的 dgm:relIds 引用 4 个（dm/lo/qs/cs） | ✓ PowerPoint 正确渲染 |

**测试 4：drawing.xml 必要性验证（test_smartart_no_drawing.py）**

| 测试 | drawing.xml | colors.xml | 文件大小 | diagram parts | PowerPoint 渲染结果 |
|------|------------|-----------|---------|--------------|-------------------|
| A | **无** | 修改原始（schemeClr→srgbClr） | 58KB | 12（省 3 个 drawing） | ✓ 正确渲染，配色为琥珀色 |
| B | **无** | 从零生成（简化版 39 个颜色） | 58KB | 12 | ✗ 全黑（styleLbl 不完整） |
| C | 有 | 从零生成（简化版 39 个颜色） | 63KB | 15 | ✗ 全黑（与 B 完全相同，drawing 无影响） |

**关键结论**：

1. **drawing.xml 不需要存储，不需要注入**：测试 A 无 drawing.xml 但 PowerPoint 正确渲染（从 layout+data+colors+quickStyle 自动重建）。测试 B/C 有无 drawing 完全相同，证明 drawing 只是缓存
2. **colors.xml 不能从零生成**：测试 B/C 全黑，因为原始 colors.xml 有 131 个颜色映射覆盖 13+ 个 styleLbl，简化版只有 39 个且 styleLbl 不全。必须存储原始 colors.xml，注入时克隆并替换 schemeClr→srgbClr
3. **SmartArt 文本只需改 data**：无 drawing 时 PowerPoint 从 data 重新渲染，自动同步。之前的发现"必须同时改 data+drawing"只在有 drawing 时成立；**不注入 drawing 时，只改 data 即可**
4. **存储 4 个 XML，注入 4 个 XML**：layout + quickStyle + data_sample + colors。省去 drawing（每个约 8-20KB），存储量减少约 40%

**Round-Trip 关键发现**：

1. **SmartArt 配色必须改 colorsN.xml**：PowerPoint 从 colors XML 重新计算渲染配色。正确做法：将 `colorsN.xml` 中的 `<a:schemeClr val="accent1"/>` 替换为 `<a:srgbClr val="D97706"/>`
2. **drawing.xml 不需要**：PowerPoint 从 layout+data+colors+quickStyle 自动重建渲染。不注入 drawing 时文本只需改 data 即可，PowerPoint 自动同步
3. **colors.xml 必须存原始版，不能从零生成**：原始文件包含 131 个颜色映射覆盖所有 styleLbl，简化版会导致全黑。注入时克隆原始 colors 并替换 schemeClr→srgbClr
4. **data XML 的 pt 节点 type 不是 "node"**：实际文本节点 type 为空字符串，需按"有 a:t 且 type ∉ {doc, parTrans, sibTrans, pres}"筛选
5. **diagram MIME 类型**：4 种需要注入的类型（data→diagramData, layout→diagramLayout, quickStyle→diagramStyle, colors→diagramColors）
6. **schemeClr 角色映射**：colors XML 中出现的角色有 accent1-6、lt1、dk1、tx1，品牌配色映射需覆盖全部

结论：**提取 → 修改 → 注入 → PowerPoint 渲染** 全链路已验证通过。SmartArt 的配色、文本均可精确控制，技术可行性完全确认。

测试脚本：`tests/test_smartart_inject.py`、`tests/test_groupshape_extract.py`、`tests/test_art_pptx.py`、`tests/test_art_smartart_deep.py`、`tests/test_smartart_roundtrip.py`、`tests/test_smartart_no_drawing.py`

> 组件库的设计方案详见独立文档：`docs/component-library-design.md`

#### 三层架构

```
┌─────────────────────────────────────────────────┐
│ ① XML 提取层                                      │
│    SmartArtExtractor / GroupExtractor / OLEExtractor │
│    输入：.pptx 的 ZIP 内 XML                      │
│    输出：结构化数据（texts, structure, style_hints）  │
└────────────────────┬────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────┐
│ ② 组件库层                                        │
│    ComponentLibrary (SQLite + 文件系统)            │
│    存储数万个预置 SmartArt/Group 模板               │
│    每个组件 = XML 模板 + 元数据 + 缩略图            │
└────────────────────┬────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────┐
│ ③ 渲染桥接层                                      │
│    ComponentRenderer                              │
│    匹配组件库 → 填入数据 → 应用品牌色/字体 → 写入 slide │
└─────────────────────────────────────────────────┘
```

#### 8.7.1 XML 提取层

**SmartArt 提取**：

.pptx 中 SmartArt 的存储结构：
```
ppt/slideN.xml          → 引用 SmartArt 的 rId（<dgm:relIds r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>）
ppt/slideN.xml.rels     → rId → "../diagrams/dataN.xml" + layoutN.xml + quickStyleN.xml + colorsN.xml
ppt/diagrams/dataN.xml       → 逻辑数据（dgm:pt / a:t 节点，文本节点的 type 属性为空字符串）
ppt/diagrams/layoutN.xml     → 布局模板（uniqueId 含具体类型名 + dgm:alg type=lin/cycle/hierarchy/...）
ppt/diagrams/colorsN.xml     → 配色方案（schemeClr 角色：accent1-6, lt1, dk1, tx1）——PowerPoint 渲染时读取此文件
ppt/diagrams/quickStyleN.xml → 快速样式（字体/效果角色映射）
ppt/diagrams/drawingN.xml    → 预渲染缓存（PowerPoint 自动重建，不需要存储/注入）
```

**重要**：修改 SmartArt 时只需处理 4 个 XML（layout + quickStyle + data + colors），不需要 drawing：
- **colorsN.xml**：克隆原始文件，将 `schemeClr` 替换为 `srgbClr`（不能从零生成，原始文件有 131 个颜色映射覆盖所有 styleLbl）
- **dataN.xml**：修改 `dgm:pt/a:t` 文本（PowerPoint 自动同步到渲染，无需改 drawing）
- **layoutN.xml** + **quickStyleN.xml**：原样注入，不修改

```python
class SmartArtExtractor:
    def extract(self, pptx_path: str, slide_index: int, shape_rId: str) -> dict:
        """
        1. zipfile 打开 .pptx
        2. 从 slideN.xml.rels 找到 diagrams/ 下的 5 个 XML 文件
        3. 解析 dataN.xml → 提取所有文本节点（dgm:pt / a:t）
           注意：文本节点的 type 属性为空字符串，不是 "node"
           需按"有 a:t 子元素且 type 不在 {doc, parTrans, sibTrans, pres} 中"筛选
        4. 解析 layoutN.xml → 提取 uniqueId（精确类型名）+ alg/@type（大类）
        5. 解析 colorsN.xml → 提取 schemeClr 角色映射（accent1-6, lt1, dk1, tx1）
           注意：colors.xml 必须存原始版（131 个颜色映射），不能从零生成（简化版会导致全黑）
        6. 返回结构化数据
        """
        return {
            "type": "smartart",
            "category": "process",      # 从 layout XML 的 alg 类型推断
            "variant": "chevron",       # 从 layout XML 的 uniqueId 精确识别
            "nodes": [
                {"id": 0, "text": "需求分析", "level": 0},
                {"id": 1, "text": "方案设计", "level": 0},
                {"id": 2, "text": "开发测试", "level": 0},
                {"id": 3, "text": "上线运维", "level": 0},
            ],
            "connections": [(0,1), (1,2), (2,3)],  # 从 layout XML 的 adj 推断
            "color_roles": {"fill": "accent1", "line": "lt1", "text": "tx1"},  # 从 colors XML 提取
            "xml_parts": {  # 4 个需要存储的 XML（drawing 不需要，PowerPoint 自动重建）
                "data": b"...", "layout": b"...", "colors": b"...",
                "quickStyle": b"...",
            },
        }
```

**GroupShape 提取**：

```python
class GroupExtractor:
    def extract(self, group_shape) -> dict:
        """
        递归遍历 GroupShape 内部元素：
        - 子文本框 → texts[]
        - 子图片 → images[]
        - 子形状（矩形/箭头/连接线） → shapes[]（记录类型+位置）
        - 嵌套 GroupShape → 递归
        """
        return {
            "type": "group",
            "texts": ["步骤一", "步骤二"],
            "images": [],
            "shapes": [
                {"type": "rectangle", "bounds": (0.5, 1.0, 3.0, 1.5)},
                {"type": "arrow", "from": (3.5, 1.75), "to": (4.0, 1.75)},
            ],
            "inferred_category": "process",  # 从形状排列推断
        }
```

**OLE/嵌入对象提取**：

```python
class OLEExtractor:
    def extract(self, pptx_path: str, slide_index: int, shape_rId: str) -> dict:
        """
        1. 从 slideN.xml.rels 找到 embedded 对象
        2. 读取 oleObject 的 progId（如 "Excel.Sheet.12"）
        3. 尝试读取关联的 chart XML（如果有）
        4. 返回元信息（类型、程序ID、是否有图表数据）
        """
        return {
            "type": "ole",
            "prog_id": "Excel.Sheet.12",
            "has_chart": True,
            "chart_data": {...},  # 如果能提取
        }
```

#### 8.7.2 组件库层

**目标**：管理数万个 SmartArt/Group 模板组件，支持高效检索、加载、编辑。

**存储方案**：SQLite 索引 + 文件系统存 XML

```
component_library/
├── index.db                          # SQLite 索引数据库
├── smartart/
│   ├── process/
│   │   ├── chevron_4node/           # SmartArt 组件目录（4 个 XML parts）
│   │   │   ├── data.xml             # 逻辑数据（dgm:pt/a:t）
│   │   │   ├── layout.xml           # 布局模板（uniqueId + alg type）
│   │   │   ├── colors.xml           # 配色方案（schemeClr 角色，注入时克隆替换）
│   │   │   └── quickStyle.xml       # 快速样式
│   │   ├── chevron_4node.png        # 缩略图
│   │   ├── chevron_4node.meta.json  # 元数据
│   │   ├── chevron_5node/
│   │   │   └── ...（同上 4 个 XML）
│   │   └── ...
│   ├── hierarchy/
│   │   ├── orgchart_3level/
│   │   │   └── ...（4 个 XML）
│   │   └── ...
│   ├── cycle/
│   ├── matrix/
│   ├── pyramid/
│   ├── relationship/
│   └── picture/
├── group/
│   ├── infographic/
│   │   ├── timeline_horizontal.xml  # GroupShape 单 XML
│   │   ├── stats_4col.xml
│   │   └── ...
│   ├── diagram/
│   │   ├── venn_3.xml
│   │   ├── funnel_5stage.xml
│   │   └── ...
│   └── layout/
│       ├── sidebar_with_stats.xml
│       └── ...
└── ole/
    ├── excel_chart_line.xml
    ├── excel_chart_bar.xml
    └── ...
```

**SQLite Schema**：

```sql
CREATE TABLE components (
    id          INTEGER PRIMARY KEY,
    type        TEXT NOT NULL,           -- 'smartart' / 'group' / 'ole'
    category    TEXT NOT NULL,           -- 'process' / 'hierarchy' / 'cycle' / ...
    variant     TEXT,                    -- 'chevron' / 'arrows' / 'orgchart' / ...
    node_count  INTEGER,                 -- 节点数量（4/5/6/...）
    level_count INTEGER,                 -- 层级数量（1/2/3/...）
    tags        TEXT,                    -- JSON array: ["flow","horizontal","arrow"]
    xml_path    TEXT NOT NULL,           -- SmartArt: 目录路径（含 4 个 XML parts）；Group/OLE: 单文件路径
    thumb_path  TEXT,                    -- 缩略图路径
    source      TEXT,                    -- 来源: 'builtin' / 'microsoft' / 'custom'
    created_at  TEXT,
    updated_at  TEXT
);

-- 全文搜索索引
CREATE VIRTUAL TABLE components_fts USING fts5(
    id, type, category, variant, tags
);

-- 高频查询索引
CREATE INDEX idx_type_category ON components(type, category);
CREATE INDEX idx_node_count ON components(node_count);
CREATE INDEX idx_type_category_nodes ON components(type, category, node_count);
```

**核心操作**：

```python
class ComponentLibrary:
    def __init__(self, db_path: str = "component_library/index.db"):
        self._db = sqlite3.connect(db_path)

    def search(self, type: str, category: str, node_count: int = None,
               tags: list[str] = None, limit: int = 10) -> list[Component]:
        """按类型/分类/节点数/标签检索组件"""

    def load_xml(self, component_id: int) -> dict:
        """加载组件的 XML 模板（4 个 parts: data/layout/colors/quickStyle，不需要 drawing）"""

    def load_thumbnail(self, component_id: int) -> bytes:
        """加载缩略图（用于方案预览）"""

    def add(self, type: str, category: str, variant: str,
            xml_parts: dict, thumb: bytes = None, **meta) -> int:
        """新增组件（SmartArt: xml_parts 含 4 个 XML parts；Group/OLE: 单 XML）"""

    def bulk_import(self, pptx_paths: list[str]) -> dict:
        """
        批量导入：从多个 PPT 中提取所有 SmartArt/Group → 入库
        返回：{"added": 1523, "skipped": 47, "errors": 12}
        """

    def match(self, extracted: dict) -> Component | None:
        """
        匹配引擎：给定提取的结构化数据，找最相似的组件
        1. 按 type + category 过滤
        2. 按 node_count 过滤
        3. 按 XML 结构相似度排序（可选，计算量大时跳过）
        4. 返回最佳匹配
        """

    def stats(self) -> dict:
        """统计信息：{"smartart": 8500, "group": 3200, "ole": 450, "total": 12150}"""
```

#### 8.7.3 渲染桥接层

**核心逻辑**：提取 → 匹配组件库 → 填入数据 → 应用品牌 → 写入 slide

```python
class ComponentRenderer:
    def render(self, slide, complex_element: dict, brand_spec: BrandSpec,
               component_lib: ComponentLibrary) -> bool:
        """
        1. 从 complex_element 获取 type/category/nodes
        2. component_lib.match() 找到最匹配的组件模板
        3. 加载组件 XML（4 个 parts: data/layout/colors/quickStyle，不需要 drawing）
        4. 将 texts[] 填入 data XML 的文本节点（PowerPoint 自动同步到渲染）
        5. 应用 brand_spec 的配色到 colors XML（克隆原始→schemeClr → srgbClr）
        6. 将处理后的 XML 写入 slide（需正确设置 rels 和 Content_Types）
        """

    def render_smartart(self, slide, element: dict, brand_spec, lib) -> bool:
        """SmartArt 专用渲染路径"""
        # 1. 匹配组件
        component = lib.match(element)
        if component is None:
            return self._fallback_diagram(slide, element, brand_spec)

        # 2. 加载模板 XML（4 个 parts，不需要 drawing）
        xml_parts = lib.load_xml(component.id)  # {data, layout, colors, quickStyle}

        # 3. 填入文本数据（只需改 data，PowerPoint 自动同步）
        filled_parts = self._fill_data(xml_parts, element["nodes"])

        # 4. 应用品牌色（克隆 colors XML → 替换 schemeClr 为 srgbClr）
        styled_parts = self._apply_brand_colors(filled_parts, brand_spec)

        # 5. 写入 slide（5 个 rId：rId1=slideLayout + rId2-5=dm/lo/qs/cs）
        self._inject_to_slide(slide, styled_parts, element["bounds"])
        return True

    def _fill_data(self, xml_parts: dict, nodes: list[dict]) -> dict:
        """将节点数据填入 data XML 的 a:t 节点。
        只需修改 data XML（dgm:pt/a:t），PowerPoint 自动同步到渲染。
        不需要修改 drawing XML（不存储、不注入）。
        """

    def _apply_brand_colors(self, xml_parts: dict, brand_spec: BrandSpec) -> dict:
        """替换 colors XML 中的配色角色为 BrandSpec 的实际颜色。
        关键：克隆原始 colors XML → 将 schemeClr 替换为 srgbClr。
        原始 colors XML 有 131 个颜色映射覆盖所有 styleLbl，不能从零生成（简化版会导致全黑）。
        brand_spec 需覆盖全部 schemeClr 角色：accent1-6, lt1, dk1, tx1。
        """

    def _inject_to_slide(self, slide, xml_parts: dict, bounds: tuple) -> None:
        """将处理后的 4 个 XML parts 写入 slide。
        需正确设置 slide rels（5 个 rId：rId1=slideLayout + rId2-5=dm/lo/qs/cs）和 Content_Types（4 种 MIME 类型）。
        """

    def _fallback_diagram(self, slide, element, brand_spec) -> bool:
        """
        组件库无匹配时，降级用 DiagramEngine 渲染：
        SmartArt process → DiagramEngine.render("process", data)
        SmartArt hierarchy → DiagramEngine.render("tree", data)
        SmartArt cycle → DiagramEngine.render("cycle", data)
        SmartArt matrix → DiagramEngine.render("swot", data)
        SmartArt pyramid → DiagramEngine.render("pyramid", data)
        """
```

#### 8.7.4 组件库管理

**批量导入流程**（从现有 PPT 文件批量入库）：

```python
# 从一个包含大量 SmartArt 的 PPT 文件批量导入
lib = ComponentLibrary()
result = lib.bulk_import(["microsoft_smartart_collection.pptx"])
# → 遍历每页每个 SmartArt → XML 提取 → 分类 → 生成缩略图 → 入库

# 从客户 PPT 提取新组件（遇到组件库没有的 SmartArt 时）
extractor = SmartArtExtractor()
element = extractor.extract("client.pptx", slide_idx=3, shape_rId="rId3")
if lib.match(element) is None:
    lib.add(type="smartart", category=element["category"],
            variant=element["variant"], xml_parts=element["xml_parts"])
```

**缩略图生成**：

```python
# 方案 1：用 python-pptx 渲染到空白 slide → 用 comtypes/win32com 转图片（Windows only）
# 方案 2：用 LibreOffice headless 转图片（跨平台）
# 方案 3：从 XML 解析出节点数/类型 → 用 Pillow 画简易示意图（最快，不依赖外部）
```

**版本管理**：

组件库随系统迭代更新，需要版本管理：

```
component_library/
├── index.db
├── versions/
│   ├── v1_schema.sql           # 初始 schema
│   ├── v2_add_level_count.sql  # 增加字段
│   └── ...
└── migrations.py               # 自动迁移
```

#### 8.7.5 与美化改版的集成

美化改版遇到 complex_elements 时的处理流程：

```
客户 PPT 中的 SmartArt
  │
  ↓ SmartArtExtractor.extract()
结构化数据 {type, category, nodes, connections}
  │
  ↓ ComponentLibrary.match()
匹配到组件库中的模板？
  │
  ├─ 是 → ComponentRenderer.render_smartart()
  │       用组件库模板 + 原数据 + 新品牌色 → 写入新 slide
  │       效果：SmartArt 保留结构，配色/字体统一
  │
  └─ 否 → ComponentRenderer._fallback_diagram()
          用 DiagramEngine 重新渲染
          效果：SmartArt 降级为 DiagramEngine 图形，结构近似但样式不同
```

#### 8.7.6 新增文件

| 文件 | 操作 | 说明 |
|------|------|------|
| `src/ppt_pro_max/enterprise/smartart_extractor.py` | **新增** | SmartArt XML 解析提取 |
| `src/ppt_pro_max/enterprise/group_extractor.py` | **新增** | GroupShape 递归提取 |
| `src/ppt_pro_max/enterprise/ole_extractor.py` | **新增** | OLE/嵌入对象提取 |
| `src/ppt_pro_max/enterprise/component_library.py` | **新增** | 组件库：SQLite 索引 + CRUD + 匹配 |
| `src/ppt_pro_max/enterprise/component_renderer.py` | **新增** | 组件渲染桥接：匹配→填数据→应用品牌→写入 |
| `src/ppt_pro_max/enterprise/component_importer.py` | **新增** | 批量导入工具：从 PPT 批量提取入库 |
| `component_library/` | **新增目录** | 组件库存储：index.db + smartart/ + group/ + ole/ |

---

## 九、修订流程（从零生成后）

用户确认方案并完整生成后，可通过对话要求修改：

- "第3页换个架构图" → 修改 page dict 的 diagram，重新 render_slide()
- "封面换个图" → 替换 image，重新 render_slide()
- "加一页团队介绍" → 新增 page，append 到 pages[]，重新生成
- "删掉第7页" → 删除 page，重新生成
- "整体换个配色" → 换 palette 原子，全量重新生成

实现：现有 `--pages` 参数 + `PageRevisionEngine`（`page_revision.py`）已支持增删改排序，但仅操作模板中的已有 slide（XML 复制），不支持内容级修改。内容级修改需配合 render_slide() 重新渲染指定页。

---

## 十、修订流程（美化改版后）

美化改版后的修改更轻量，因为内容已提取为 pages[]：

- "第3页换个配色" → 换 palette 原子，重新 render_slide(第3页)
- "封面图换一张" → 替换 page.image，重新 render_slide(第1页)
- "整体字体换大一号" → 调 density，全量重新渲染
- "第5页删掉那个图" → page.image = None，重新 render_slide(第5页)

与从零生成的区别：内容固定不变，只调整视觉参数。

---

## 附录：文档修订记录

以下为本次审查发现并修正的问题清单：

| # | 原文问题 | 修正 |
|---|----------|------|
| 1 | 问题表缺少"代码现状"列，无法验证问题是否真实存在 | 新增"代码现状"列，标注具体文件和行号 |
| 2 | 问题1"Pipeline 渲染用 paragraph 级字体"描述不准确 | `_populate_slide()` 的 title/bullets 已用 run-level；paragraph 级实际在 cards（1010-1021）、code（1148-1162）、exercise（1208-1227） |
| 3 | 问题2"图片直接 add_picture 拉伸"描述不准确 | `_insert_content_image()` 保持宽高比不变形（用 Pillow 读尺寸），问题是留白/不填满；Hero 图片路径已有 Pillow 裁剪 |
| 4 | 问题4"有模板时走 EnterpriseRenderer（低质量）"描述过于笼统 | 补充具体代码位置 `pipeline.py:139` 的 `use_freestyle_render` 逻辑 |
| 5 | 问题5"AI 配图需手动配置 image_prompt"部分不准确 | `ImageFetcher._build_image_prompt()` 已能自动生成 prompt，问题在于 Pipeline 未在缺图时自动调用 |
| 6 | 3.1 自动模板分发逻辑使用的模板名不存在 | `numbered_list`/`process_flow`/`3cards`/`4cards`/`table_of_contents`/`bullet_list` 在 `layout_registry.py` 中不存在，替换为实际存在的 key |
| 7 | 3.1 page dict 缺少 `chart`/`notes`/`links` 字段 | Pipeline 的 `_build_page_designs()` 已输出这些字段，补充 |
| 8 | 3.2 ContentParser 逻辑与现有代码重复 | 标注 `load_enterprise_content()` 已实现 content.json 解析，仅需扩展 README 解析 |
| 9 | 3.2 StoryPlanner 策略中无 `toc` goal | 标注需在 StoryPlanner 增加或映射 |
| 10 | 3.3 图片优先逻辑有两个编号"5" | 修正为 5 和 6 |
| 11 | 3.3 未说明现有 `image_matcher.py` 和 `scanner.py` 已有部分能力 | 补充现有代码说明 |
| 12 | 3.4 方案预览写"3 页"与"4 页"矛盾 | 统一为 4 页，目录页替换为 CTA 页 |
| 13 | 3.5 Pipeline 修订未说明第三条渲染路径 | 实际有四条渲染路径：`_populate_slide()`、`_render_with_precision_renderer()`、`_render_with_ppt_renderer()`（死代码）、`_populate_slide_with_layout()`（死代码）。修订后删除 3 条旧路径 |
| 14 | 3.5 有模板时设计为"填 placeholder"思路有误 | 用户给模板是借设计风格，不是让内容迁就 placeholder 框。改为：有模板和无模板内容渲染逻辑完全一致，仅 `create_presentation()` 来源不同 |
| 15 | 四、文件清单将 `content_parser.py` 标为"新增 content_intelligence.py" | 修正为修改现有 `content_parser.py` |
| 16 | 四、文件清单将 `image_matcher.py` 标为"新增 image_planner.py" | 修正为修改现有 `image_matcher.py` |
| 17 | 四、文件清单 `proposal_generator.py` 说明写"×3 页" | 统一为 4 页 |
| 18 | 五、API 变更未说明 `materials_dir` 与 `project` 参数的关系 | 补充说明 |
| 19 | 七、修订流程标"无需新代码"不准确 | `PageRevisionEngine` 仅做 XML 复制，不支持内容级修改，需配合 render_slide() |
| 20 | 架构图模块名与实际代码不一致 | ContentParser 标注为扩展现有 `content_parser.py`，ProposalGenerator 标注为新增 |
| 21 | 新增需求：客户 PPT 美化改版 | 新增第八节（美化改版模式），含 SlideExtractor、beautify 渲染策略、API 设计、实施顺序 P9-P11 |
| 22 | SmartArt/GroupShape 标为"无法提取"是错误判断 | .pptx 是 ZIP 包，XML 全在里面，lxml 直接解析即可。新增 8.7 节：三层架构（XML 提取层→组件库层→渲染桥接层），支持万级组件管理 |
| 23 | 组件库需求：数万个 SmartArt/Group 模板的高效管理 | 新增 SQLite 索引 + 文件系统存储方案，含 Schema、全文搜索、批量导入、匹配引擎、版本管理 |
| 24 | 8.7.1 SmartArt 存储结构使用 ppt/dgm/ 路径错误 | 实际路径是 `ppt/diagrams/`，且无中间的 smartArtN.xml 中转文件，slide rels 直接引用 5 个 XML |
| 25 | SmartArt 配色改 drawing XML 无效 | PowerPoint 渲染时从 colorsN.xml 重新计算，drawing 只是缓存。正确做法：将 colorsN.xml 的 schemeClr 替换为 srgbClr |
| 26 | SmartArt 文本只改 data XML 不够 | drawingN.xml 包含预渲染文本副本，改 data 时需同步改 drawing。**已修正**：见 #29/#31，不注入 drawing 时只需改 data 即可 |
| 27 | data XML 的 pt 节点 type="node" 不存在 | 实际 SmartArt 文本节点 type 为空字符串，需按"有 a:t 且 type ∉ {doc, parTrans, sibTrans, pres}"筛选 |
| 28 | Content_Types 的 diagram MIME 类型不完整 | 5 种不同类型：diagramData/diagramLayout/diagramStyle/diagramColors/diagramDrawing，必须精确匹配 |
| 29 | drawing.xml 需要存储和注入 | **不需要**。测试 A（无 drawing）PowerPoint 正确渲染，测试 B/C（有无 drawing）完全相同。PowerPoint 从 layout+data+colors+quickStyle 自动重建。存储量减少约 40% |
| 30 | colors.xml 可以从零生成 | **不可以**。测试 B/C 从零生成的简化版（39 个颜色）导致全黑，原始版有 131 个颜色映射覆盖所有 styleLbl。必须存储原始 colors，注入时克隆并替换 schemeClr→srgbClr |
| 31 | 文本必须同时改 data + drawing | **只在有 drawing 时成立**。不注入 drawing 时，只需改 data，PowerPoint 自动同步到渲染 |
| 32 | 问题1行号引用错误 | `enterprise_renderer.py:181-183` 和 `pipeline.py:393-394` 非主要问题位置。实际 paragraph 级字体在 `pipeline.py:1010-1021`（cards）、`1148-1162`（code）、`1208-1227`（exercise）；`_populate_slide()` 的 title/bullets 已用 run-level |
| 33 | 问题2"图片直接 add_picture 拉伸" | `_insert_content_image()` 保持宽高比不变形（用 Pillow 读尺寸），问题是留白/不填满；Hero 图片路径已有 Pillow 裁剪 |
| 34 | 问题5"Pipeline 未在缺图时自动调用" | 根本原因是 `image_mode` 默认 `"placeholder"` 导致 ImageFetcher 不构建（返回 None），不仅仅是"未调用" |
| 35 | 3.1 goal→layout 映射错误 | `toc` 不存在（sidebar-left 映射 navigation/agenda/overview）；`market` 映射到 content-with-title 非 funnel；`code_demo` 应为 `code`；content-with-title 映射 10 个 goal 非仅 problem |
| 36 | 3.5 渲染路径数量 | 实际 4 条非 3 条：遗漏 `_populate_slide_with_layout()`（pipeline.py:652-772）。4 条路径中仅 `_render_with_precision_renderer()` 和 `_populate_slide()` 被调用 |
| 37 | density_profile 描述"仅限字号" | 实际调节 8 个字段：title_size/subtitle_size/body_size/bullet_size/line_spacing/max_bullets/max_bullet_chars/image_width_ratio |

---

## 附录 B：实施效果预估

### 逐阶段效果评估

| 阶段 | 改动量 | 效果确定性 | 预期效果 | 风险 |
|------|--------|-----------|---------|------|
| **P1** render_slide() | 中（~200 行） | **高** | 有模板场景字体不再回退宋体，图片不再变形 | 低。PrecisionRenderer 的 add_text/add_image/add_card 已验证可用，render_slide() 本质是组合调用 |
| **P2** Pipeline 统一走 PrecisionRenderer | 小（删代码为主） | **高** | 消除双路径质量差异 | 中。需回归测试现有 30 个测试用例，`_populate_slide()` 删除后 `test_populate_slide.py` 需重写 |
| **P3** mood_words 扩展 | 小（加几行映射） | **高** | "international/cream/frosted" 等风格词能被识别 | 极低。纯数据扩展 |
| **P4** ContentParser README 解析 | 中（~150 行） | **中** | 有 README.md 的项目能自动生成 pages[] | 中。Markdown 结构千差万别，标题层级、代码块、表格的解析边界情况多；纯中文 README 的 goal 推断依赖关键词匹配，准确率约 70-80% |
| **P5** 图片尺寸分类 + image_prompt 自动生成 | 小（~80 行） | **中高** | 缺图页面自动补 AI 配图 | 中低。尺寸分类逻辑简单；image_prompt 自动生成依赖 ImageFetcher 已有逻辑，但 AI 生图质量本身不可控（可能生成与内容不相关的图） |
| **P6** 方案预览 ProposalGenerator | 中（~200 行） | **中** | 用户出完整 PPT 前能看 2-3 个风格方案 | 中。预览 PPT 的 4 页内容是"假数据"（用示例 bullets/cards），用户可能误以为这是最终内容；方案差异化策略（同 mood 换 palette）可能产生视觉差异不够明显的方案 |
| **P7** generate_ppt() API 集成 | 小（~50 行） | **高** | proposal/confirmed_proposal/materials_dir 参数可用 | 低。API 层改动简单 |
| **P9** SlideExtractor + beautify 渲染 | 大（~300 行） | **中高** | 客户 PPT 能提取内容并美化 | 中。简单文本 PPT 效果确定；复杂元素走降级路径 |
| **P10** beautify API + CLI | 小（~40 行） | **高** | --beautify 参数可用 | 低 |
| **P12** SmartArt/Group XML 提取器 | 中（~250 行） | **高** | SmartArt/GroupShape 文本和结构可提取 | **极低**。Round-trip 测试已验证：提取→修改配色→修改文本→注入新PPT→PowerPoint正确渲染，全链路通过 |
| **P13** 组件库基础设施 | 中（~300 行） | **高** | SQLite 索引 + CRUD + 批量导入可用 | 低。SQLite + 文件系统是成熟方案 |
| **P14** 组件库匹配 + 渲染桥接 | 中（~200 行） | **中高** | SmartArt 可用组件库模板重建 | **低**。Round-trip 已验证 SmartArt 注入+渲染可行，渲染桥接本质是组合调用 |

### P1-P2（核心修复）效果：确定性高

这是整个方案价值最大的部分。当前代码中 PrecisionRenderer 的 run-level 字体、Pillow 裁剪、品牌装饰（accent 竖条/muted 底条/LOGO/页脚/水印）**已经实现且可用**，问题仅在于 Pipeline 有模板时不走这条路径。P2 本质是删掉低质量路径，让已有高质量代码接管。

**预期**：有模板的企业场景输出质量从"不可用"提升到与无模板场景一致。这是确定性的改善。

### P4（README 解析）效果：有限

README.md 的结构没有统一规范。实际场景中：
- 规范的 README（一级标题分节、二级标题列点）→ 解析效果好
- 不规范的 README（混合 Markdown、大段散文、无标题结构）→ 解析结果差，需用户手动调整
- 纯中文 README → goal 推断依赖关键词，"痛点/方案/功能"能匹配，但"核心优势/产品理念"等模糊表述会归入 content

**预期**：约 60-70% 的 README 能生成可用的 pages[]，其余需用户补充 content.json。这比"必须手写 content.json"已是进步，但不是银弹。

### P6（方案预览）效果：中等

方案预览的核心价值是让用户在生成 14 页完整 PPT 之前确认风格方向。但有两个现实问题：

1. **预览内容是假的**：4 页预览用的是示例数据（"要点1/要点2"），不是用户真实内容。用户确认的是"配色+字体+装饰"组合，不是内容布局。这其实够用——方案预览的目的就是确认风格，不是确认内容。

2. **方案差异可能不够明显**：当前 ThemeComposer 的差异化策略是换 palette（如 ocean-blue → golden-luxury），但同一 mood 下的 palette 可能在 PowerPoint 中视觉差异不大（尤其是浅色系）。dark 模式下差异更明显。

**预期**：方案预览能有效避免"做完 14 页发现风格不对要重来"的问题，但用户可能需要 2-3 轮预览才能找到满意方案。

### P9（美化改版）效果：中等偏高

之前评估为"中低"是因为认为 SmartArt/GroupShape 无法提取。实际上 .pptx 是 ZIP 包，所有数据都在 XML 里，lxml 直接解析即可突破 python-pptx 高层 API 的限制。

**剩余的实际难题**：

**难题 1：布局还原度**

`layout_hint` 记录原始位置，但 PrecisionRenderer 的 render_slide() 是按 goal→layout 分发的，不是按坐标自由摆放。美化后的布局会"规范化"——标题统一在左上、正文统一在左侧、图片统一在右侧。这对设计一致性是好事，但意味着美化后的 PPT 与原 PPT 的布局会有明显差异，客户可能觉得"不像我的 PPT 了"。

**难题 2：内容量不匹配**

客户 PPT 一页可能塞了 10 条 bullets + 3 张图，而 PrecisionRenderer 的 density_profile 最多支持 15 条（density=10）。如果客户 PPT 内容密度超出系统上限，要么截断内容，要么字号极小，效果都不好。

**难题 3：组件库冷启动**

组件库初期为空，所有 SmartArt 都走 DiagramEngine 降级路径。效果可用但样式不同。随着批量导入（P12-P13），覆盖率逐步提升，匹配成功率从 0% 增长到 70-80%+。

**预期效果分级**：

| 客户 PPT 类型 | 提取成功率 | 美化效果 | 说明 |
|--------------|-----------|---------|------|
| 简单文本 PPT（标题+正文+少量图片） | 85-90% | **好** | 最适合美化，效果提升明显 |
| 中等复杂度（含 SmartArt/图表/表格） | 75-85% | **中好** | SmartArt 通过 XML 提取+组件库/降级路径可处理 |
| 高复杂度（大量 GroupShape/嵌入对象/动画） | 60-70% | **中** | GroupShape 可提取但重建依赖组件库覆盖度；嵌入 Excel 图表数据提取受限 |

### 总体预期

| 维度 | 当前 | P1-P3 后 | P1-P8 后 | P1-P14 后 |
|------|------|----------|----------|-----------|
| 有模板场景输出质量 | 不可用（宋体+变形图） | **可用** | **可用**（已验证8种goal类型） | 可用 |
| 无模板场景输出质量 | 可用 | 可用 | **可用**（已验证6-8页） | 可用 |
| 用户需手写 content.json | 必须 | 必须 | **可选**（README 可替代） | 可选 |
| 风格确认机制 | 无（一次出完） | 无 | **有**（方案预览，3方案×4页） | 有 |
| AI 配图 | 需手动配置 | 需手动配置 | **自动**（缺图时补prompt） | 自动 |
| 美化改版 | 不支持 | 不支持 | 不支持 | **支持**（含 SmartArt/Group） |
| SmartArt/Group 提取 | 不支持 | 不支持 | 不支持 | **支持**（XML 解析） |
| 组件库 | 无 | 无 | 无 | **有**（SQLite 索引，支持万级组件） |

### P1-P8 端到端评估结果

**545 tests passed, 5 skipped, 0 failures. Lint clean on all modified files.**

| ID | 场景 | 页数 | 形状数 | 图片数 | 最小字号 | 最大字号 | 文件大小 |
|----|------|------|--------|--------|----------|----------|----------|
| E1 | Freestyle dark cyberpunk | 6 | 33 | 4 | 16pt | 44pt | 52.8 KB |
| E2 | Freestyle warm elegant | 8 | 43 | 5 | 14pt | 44pt | 59.9 KB |
| E3 | Enterprise 全量（模板+品牌+内容+图片） | 8 | 57 | 7 | 11pt | 52pt | 39.5 KB |
| E4 | Enterprise 仅 README.md | 6 | 31 | 6 | 11pt | 52pt | 43.4 KB |
| E5A | 方案A（mckinsey） | 4 | 27 | 0 | 11pt | 52pt | 31.8 KB |
| E5B | 方案B（alt palette） | 4 | 27 | 0 | 11pt | 52pt | 31.8 KB |
| E5C | 方案C（alt mood） | 4 | 27 | 0 | 11pt | 52pt | 31.8 KB |

#### E3 内容验证（8种 goal 类型全部正确渲染）

| 页面 | Goal | 渲染内容 |
|------|------|----------|
| 0 | hook | 标题 + 副标题 + hero 图片 |
| 1 | problem | 标题 + 4条 bullets + 图片 |
| 2 | solution | 标题 + 4条 bullets + 图片 |
| 3 | features | 标题 + 3张卡片（AI Engine, Live Dashboard, Integration） |
| 4 | data | 标题 + 表格图（5行） |
| 5 | code | 标题 + 代码块（python, 4行） |
| 6 | exercise | 标题 + 徽章 "Exercise 5 min" + 3步骤 |
| 7 | cta | 标题 + 副标题 |

#### 质量检查

- 所有字号 >= 11pt（无不可读文字）
- 所有页面 >= 3 个形状（无空白页）
- 所有页面有文字内容（无空页）
- 图片正确放置：hero→hook, product→features
- README.md 正确解析为 6 页，含 goal 推断
- 方案差异化：A=indigo-deep, B=slate-minimal, C=lavender-dream

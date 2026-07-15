# Enterprise Pipeline 设计方案

> 版本: v3.0 | 日期: 2026-07-11 | 状态: Phase A-H 全部完成，375 测试通过，lint clean

## 1. 背景与目标

### 1.1 现状

当前 PPT Pro Max 采用 FreeStyle Pipeline（StoryPlanner → DesignDecider → ContentGenerator → PPTRenderer），从空白 Presentation 出发自由发挥。适用于个人/创意场景，但无法满足企业客户需求：

- 无法读取企业 PPT 模板，颜色/字体/排版必然偏离品牌规范
- 无法插入 LOGO
- 无法匹配产品图片
- 无用户确认环节，生成结果不可控
- 布局信息密度低，无法承载教育/培训等高信息量场景

### 1.2 目标

在复用现有基础组件的前提下，新增 Enterprise Pipeline，适配企业客户的复杂业务场景：

1. **模板驱动**：基于企业 PPT 模板渲染，遵循品牌规范
2. **资产整合**：用户只需提供一个项目文件夹，系统自动识别资产
3. **确认机制**：生成前展示方案，用户确认后再执行
4. **高密度支持**：教育/培训场景的大信息量 PPT
5. **零侵入**：FreeStyle Pipeline 完全不动，`--project` 不传时走原有逻辑

### 1.3 核心原则

- **有什么用什么**：资产不全不报错，缺什么用默认逻辑兜底
- **复用不推翻**：共享基础组件层，不修改现有 Pipeline
- **降级优雅**：Enterprise Pipeline 内部也有降级，最差情况等价于 FreeStyle + Review Gate

---

## 2. 双管道架构

```
                    ┌──────────────────────┐
                    │   generate_ppt()     │
                    │   CLI / API 入口     │
                    └──────────┬───────────┘
                               │
                    ┌──────────▼───────────┐
                    │  有 --project ?      │
                    └──┬─────────────────┬─┘
                       │ Yes             │ No
                ┌──────▼──────┐   ┌──────▼──────┐
                │ Enterprise  │   │  FreeStyle  │
                │ Pipeline    │   │  Pipeline   │
                │ (新增)      │   │  (不动)     │
                └──────┬──────┘   └─────────────┘
                       │
          ┌────────────▼────────────┐
          │  共享基础组件层          │
          │  ThemeComposer          │
          │  LayoutRegistry         │
          │  ImageFetcher           │
          │  ChartBuilder           │
          │  Effects                │
          │  QAGates                │
          │  StoryPlanner           │
          └─────────────────────────┘
```

### 2.1 分流逻辑

```python
def generate_ppt(query, **kwargs):
    project_dir = kwargs.get('project')
    if project_dir:
        return EnterprisePipeline.run(query, project_dir, **kwargs)
    else:
        return _generate_ppt_freestyle(query, **kwargs)  # 现有逻辑，零改动
```

> **实现说明**：当前 `generate_ppt()` 的 FreeStyle 逻辑是内联的（见 `__init__.py`），需重构为 `_generate_ppt_freestyle()` 独立函数，再在 `generate_ppt()` 中根据 `--project` 分流。新增参数（`project`、`business_mode`、`review`、`review_file`、`output_version`、`from_version`、`history`）需添加到函数签名中。

### 2.2 共享组件复用关系

| 组件 | FreeStyle | Enterprise | 复用方式 |
|------|-----------|------------|---------|
| StoryPlanner | 直接使用 | 直接使用 | 不动 |
| ThemeComposer | 直接使用 | 无模板时兜底 | 不动 |
| LayoutRegistry | 直接使用 | 参考坐标 + 密集布局扩展 | 扩展 |
| ImageFetcher | 直接使用 | 图片池匹配不上时兜底 | 不动 |
| ChartBuilder | 直接使用 | 直接使用 | 不动 |
| DiagramEngine | 不存在 | 直接使用 | 新建，共享 |
| Effects | 直接使用 | 直接使用 | 不动 |
| QAGates | 直接使用 | 直接使用 | 不动 |
| ThemeMapper | 直接使用 | 扩展使用（BrandSpec 映射） | 扩展 |
| ContentGenerator | 直接使用 | 复用 + 图片池匹配扩展 | 扩展 |
| DesignDecider | 直接使用 | 不使用，新建受约束版本 | 不动 |
| PPTRenderer | 直接使用 | 不使用，新建模板渲染版本 | 不动 |
| PageRevisionEngine | 不存在 | 直接使用 | 新建，Enterprise 专用 |
| ImageMatcher | 不存在 | 直接使用 | 新建，Enterprise 专用 |

---

## 3. Enterprise Pipeline 流程

```
ProjectScanner → TemplateAnalyzer → StoryPlanner → EnterpriseDesignDecider → ReviewGate → ContentGenerator → EnterpriseRenderer
   (新增)          (新增)            (复用)          (新增)                 (新增)      (复用+扩展)       (新增)
```

### 3.1 Phase 0: ProjectScanner（资产扫描）

**职责**：扫描项目文件夹，识别所有可用资产。

**项目文件夹结构**（扁平，无子目录要求）：

```
my_project/
├── template.pptx      # 有就用，没有就自由发挥
├── logo.png           # 有就插，没有就跳过
├── brand.json         # 有就读，没有就从模板提取
├── content.json       # 有就覆盖，没有就自动生成
├── hero.png           # 其余图片全部收入图片池
├── dashboard.png
├── team.jpg
└── output/            # ★ 输出目录（自动创建）
    ├── v1/            # 第1版
    │   ├── presentation.pptx
    │   └── meta.json  # 版本元数据
    ├── v2/            # 第2版（修订后）
    │   ├── presentation.pptx
    │   └── meta.json
    └── latest -> v2/  # 符号链接指向最新版（Windows 用 junction）
```

**识别规则**：

| 资产类型 | 识别方式 | 必需 |
|---------|---------|------|
| PPT 模板 | 项目根目录下（不递归，排除 output/ 子目录）唯一的 `.pptx` 文件 | 否 |
| LOGO | 文件名（去掉扩展名后）等于 `logo`，或以 `logo` 开头后跟分隔符（`_`/`-`），或以 `logo` 结尾 | 否 |
| 品牌规范 | `brand.json` | 否 |
| 内容数据 | `content.json` | 否 |
| 图片池 | 其余所有图片文件（png/jpg/jpeg/svg/webp） | 否 |

**输出数据结构**：

```python
@dataclass
class ProjectAsset:
    project_dir: str
    template_path: Optional[str] = None
    logo_path: Optional[str] = None
    brand_raw: Optional[dict] = None       # brand.json 原始内容
    content_raw: Optional[dict] = None     # content.json 原始内容
    image_pool: List[str] = field(default_factory=list)  # 图片文件路径列表
```

### 3.2 Phase 0.5: TemplateAnalyzer（模板解析）

**职责**：解析 template.pptx，提取品牌规范和布局信息。无模板时跳过。

**提取内容**：

| 提取项 | 来源 | 用途 | 备注 |
|--------|------|------|------|
| 主题色 | theme part XML `a:clrScheme` | 品牌颜色 | 需 XML 解析，python-pptx 无 Python API |
| 字体方案 | theme part XML `a:fontScheme` | 标题/正文字体 | 需 XML 解析，python-pptx 无 Python API |
| 颜色映射 | `slide_master` XML `p:clrMap` | 颜色角色→主题色索引 | 需 XML 解析 |
| 布局列表 | `slide_master.slide_layouts[]` | 布局映射 | python-pptx API 可用 |
| 占位符信息 | 每个 layout 的 placeholders | 内容放置位置 | `placeholder_format.type` 返回 PP_PLACEHOLDER 枚举 |
| LOGO 推断 | slide_master 上的非 placeholder 图片 shape | LOGO 位置/尺寸 | 遍历 shapes 按 shape_type 过滤 |
| 背景样式 | slide_master.background | 背景处理 | python-pptx API 可用 |

> **实现约束**：python-pptx 1.0.2 的 `SlideMaster` 对象没有 `theme` 属性（`dir(slide_master)` 中不存在）。主题色/字体方案存储在 OOXML 包的 `theme/theme1.xml` 中，必须通过 XML 解析提取：
>
> ```python
> from lxml import etree
> from pptx.opc.constants import RELATIONSHIP_TYPE as RT
> from pptx.oxml.ns import qn
>
> # 方案1: 通过 Presentation part 获取 theme part（推荐）
> # 注意：theme 关系在 Presentation part 上，不在 Package 上
> # slide_master.part.package.part_related_by(RT.THEME) 会 KeyError
> theme_part = prs.part.part_related_by(RT.THEME)
> # theme_part 是 Part 对象，没有 .element 属性，必须用 blob 解析
> theme_element = etree.fromstring(theme_part.blob)
>
> # 方案2: 通过 slide_master 的 rels 获取 theme part
> for key, rel in slide_master.part.rels.items():
>     if rel.reltype == RT.THEME:
>         theme_part = rel.target_part
>         break
> theme_element = etree.fromstring(theme_part.blob)
>
> # 从 theme_element 解析颜色和字体
> # 注意：a:clrScheme 和 a:fontScheme 是 a:themeElements 的子元素，不是 a:theme 的直接子元素
> clr_scheme = theme_element.find(qn('a:themeElements')).find(qn('a:clrScheme'))
> font_scheme = theme_element.find(qn('a:themeElements')).find(qn('a:fontScheme'))
>
> # 颜色提取：部分条目使用 <a:sysClr> 而非 <a:srgbClr>，需要 fallback
> # dk1: <a:sysClr val="windowText"/> → fallback "#000000"
> # lt1: <a:sysClr val="window"/> → fallback "#FFFFFF"
> # accent1-6, dk2, lt2: 通常使用 <a:srgbClr val="1F497D"/>
> SYS_CLR_FALLBACK = {
>     "windowText": "#000000", "window": "#FFFFFF",
>     "WindowText": "#000000", "Window": "#FFFFFF",
> }
>
> def extract_color(clr_element):
>     srgb = clr_element.find(qn('a:srgbClr'))
>     if srgb is not None:
>         return f"#{srgb.get('val')}"
>     sys_clr = clr_element.find(qn('a:sysClr'))
>     if sys_clr is not None:
>         return SYS_CLR_FALLBACK.get(sys_clr.get('val'), "#000000")
>     return "#000000"
>
> # 方案3: 从 p:clrMap 映射获取颜色角色（仅获取角色映射，不含具体 RGB 值）
> clr_map = slide_master.element.find(qn('p:clrMap'))
> # clr_map 的属性值映射到 theme color indices（如 bg1="lt1", tx1="dk1", accent1="accent1"）
> # 需结合 a:clrScheme 才能解析出具体 RGB 值
> ```

**输出数据结构**：

```python
@dataclass
class BrandSpec:
    # 来源标记
    source: str = "none"  # "template" | "brand_json" | "merged" | "none"

    # 颜色（键名与 ThemeMapper 颜色角色一致）
    colors: Optional[dict] = None
    # {"primary": "#1A3C6E", "on-primary": "#FFFFFF", "secondary": "#64748B",
    #  "accent": "#2ECC71", "background": "#FFFFFF", "foreground": "#333333",
    #  "muted": "#F1F5F9", "muted-foreground": "#666666", "border": "#E2E8F0",
    #  "destructive": "#EF4444"}

    # 字体
    fonts: Optional[dict] = None
    # {"heading": "Arial", "body": "Calibri"}

    # 间距规范
    spacing: Optional[dict] = None
    # {"title_size_pt": 28, "body_size_pt": 14, "line_spacing": 1.2,
    #  "margins_inches": 0.8, "bullet_spacing_pt": 8}

    # LOGO
    logo: Optional[dict] = None
    # {"position": "top_right", "left": 11.5, "top": 0.3,
    #  "width": 1.2, "height": 0.5, "on_every_slide": True,
    #  "skip_slides": ["title-slide", "cta-closing"]}

    # 布局映射
    layout_mapping: Optional[dict] = None
    # {"hook": 0, "content": 2, "features": 3, "metrics": 4, "cta": 5}
    # key = goal 名称, value = slide_layouts 的 index

    # 模板布局详情（从模板提取）
    template_layouts: Optional[List[dict]] = None
    # [{"index": 0, "name": "Title Slide", "placeholders": [...]}]

    # 暗色模式
    dark_mode: bool = False
```

**合并优先级**：

```
brand.json 显式声明 → 最高优先级
template.pptx 提取  → 中等优先级
ThemeComposer 兜底  → 最低优先级
```

### 3.3 Phase 1: StoryPlanner（叙事规划）

**复用现有 StoryPlanner**，新增教育策略：

| 策略 | 页数 | 结构 | 适用场景 |
|------|------|------|---------|
| YC Seed Deck | 10 | hook→problem→solution→...→cta | 创业融资（现有） |
| Product Demo | 8 | hook→problem→solution→...→cta | 产品演示（现有） |
| Sales Pitch | 7 | hook→problem→agitation→...→cta | 销售推介（现有） |
| **Education Lecture** | 15-30 | outline→concept→example→...→summary→exercise | 教育培训（新增） |
| **Corporate Training** | 12-20 | objective→concept→practice→...→assessment | 企业内训（新增） |

**Education Lecture 结构**：

```
outline → concept_1 → example_1 → concept_2 → example_2 → ... → summary → exercise → resources
```

**Corporate Training 结构**：

```
objective → agenda → concept_1 → demo_1 → practice_1 → concept_2 → ... → key_takeaways → assessment → next_steps
```

**情绪弧线**：

- Education: 引导→理解→巩固→应用→总结
- Training: 目标→认知→实践→掌握→行动

### 3.4 Phase 2: EnterpriseDesignDecider（受约束的设计决策）

**职责**：基于 BrandSpec 做设计决策，而非自由发挥。

**与 FreeStyle DesignDecider 的区别**：

| 决策项 | FreeStyle | Enterprise |
|--------|-----------|------------|
| 颜色 | ThemeComposer 自由组合 | BrandSpec.colors |
| 字体 | ThemeComposer 自由组合 | BrandSpec.fonts |
| 布局 | LayoutRegistry 自由选择 | BrandSpec.layout_mapping 映射到模板布局 |
| 装饰 | ThemeComposer decoration | 模板自带装饰，不额外添加 |
| 动画 | goal-based 选择 | 保守选择（企业场景偏克制） |
| 密度 | 参数存在但效果有限 | 根据 business_mode 和 density 参数调整 |

**布局映射逻辑**：

```
1. BrandSpec.layout_mapping 有显式映射 → 直接使用
2. 无显式映射 → 自动匹配：
   a. 按布局名称模糊匹配（"Title" → hook, "Content" → content, "Two Column" → comparison）
   b. 按占位符类型/数量匹配（有3个card位 → features）
   c. 都匹配不上 → 使用模板的空白布局 + 手动放置
```

**business_mode 参数**：

Enterprise Pipeline 通过 `business_mode` 区分业务场景，影响设计决策：

| business_mode | 密度倾向 | 布局偏好 | 动画风格 | 典型客户 |
|---------------|---------|---------|---------|---------|
| `pitch` | 中 | 现有布局 | 适度 | 融资/销售 |
| `education` | 高 | 密集布局 | 克制/教学序列 | 教育机构 |
| `training` | 高 | 密集布局+练习 | 克制 | 企业内训 |
| `report` | 高 | 表格/图表 | 极简 | 数据报告 |

### 3.5 Phase 2.5: ReviewGate（用户确认）

**职责**：展示中间方案，等用户确认后再继续。

**输出内容**：

```
═══════════════════════════════════════════
  PPT 设计方案
═══════════════════════════════════════════

📋 资产识别结果
  ✓ 模板: template.pptx (6个布局)
  ✓ LOGO: logo.png (将插入每页右上角)
  ✗ 品牌规范: 未提供 (从模板提取颜色/字体)
  ✓ 内容: content.json (8页内容)
  ✓ 图片池: 3张 (hero.png, dashboard.png, team.jpg)

🎨 设计规范
  主色: #1A3C6E (从模板提取)
  字体: Arial / Calibri (从模板提取)
  暗色模式: 否

📑 页面规划 (8页)
  1. hook        → Title Slide    → "颠覆传统招聘方式"
  2. problem     → Content        → "传统招聘的三大痛点"
  3. solution    → Content        → "AI驱动的智能匹配"
  4. product     → Image+Text     → 产品展示 (使用 hero.png)
  5. features    → Three Cards    → "核心功能" (使用 dashboard.png)
  6. traction    → Four Metrics   → "增长数据"
  7. team        → Image+Text     → (使用 team.jpg)
  8. cta         → CTA Closing    → "立即体验"

⚙️ 需要确认
  - LOGO插入每页右上角，跳过封面/结尾页，是否OK？
  - 图片池中 hero.png 匹配到产品页，是否正确？
  - 第5页使用模板的"Slide Layout 4"，是否OK？

请确认方案 (y/n/edit):
═══════════════════════════════════════════
```

**交互方式**：

| 运行模式 | ReviewGate 行为 |
|---------|----------------|
| CLI `--review` | 输出方案到终端，等待用户输入 y/n |
| CLI `--review --review-file proposal.json` | 输出方案到 JSON 文件，用户编辑后重新运行 |
| API `review=True` | 返回方案 dict，不执行渲染；再次调用时传 `proposal_id` 继续 |
| 不传 `--review` | 跳过确认，直接生成（默认行为） |

### 3.6 Phase 3: ContentGenerator（内容生成 + 图片匹配）

**复用现有 ContentGenerator**，扩展图片匹配逻辑。

**图片匹配优先级**：

```
1. content.json 中显式指定 image 字段 → 最高优先级
2. 图片池文件名模糊匹配 → 中等优先级
3. ImageFetcher 搜索/生成 → 兜底
```

**文件名模糊匹配算法**：

```python
def match_image(image_pool, image_keywords, goal):
    """
    1. 将 image_keywords 分词
    2. 对图片池中每个文件名（去掉扩展名）做分词
    3. 计算关键词重叠度
    4. 返回匹配度最高的图片，低于阈值则返回 None
    """
    scored = []
    for path in image_pool:
        filename = Path(path).stem.lower().replace('_', ' ').replace('-', ' ')
        score = keyword_overlap(image_keywords, filename)
        scored.append((path, score))
    scored.sort(key=lambda x: x[1], reverse=True)
    if scored and scored[0][1] >= THRESHOLD:
        return scored[0][0]
    return None
```

### 3.7 Phase 4: EnterpriseRenderer（模板渲染）

**职责**：基于模板渲染 PPT，而非从空白创建。

**与 FreeStyleRenderer 的核心区别**：

| | FreeStyleRenderer | EnterpriseRenderer |
|---|---|---|
| 创建方式 | `Presentation()` | `Presentation(template_path)`，模板无效时降级为 `Presentation()` |
| 布局选择 | `slide_layouts[6]` 空白布局（仅默认模板有效） | `slide_layouts[mapped_index]` 模板布局 |
| 内容放置 | 全部手动 shape | 优先填模板占位符，不足再手动 shape |
| LOGO | 无 | 自动插入（从模板 master 继承或手动插入） |
| 颜色/字体 | ThemeComposer 决定 | BrandSpec 决定 |
| 密度调整 | 无 | 根据 density 参数调整字号/间距 |
| 页面级修订 | 不支持 | 支持，委托 PageRevisionEngine 两阶段重建（见 §9.6.3） |

**EnterpriseRenderer 创建 Presentation 逻辑**：

```python
def _create_presentation(self, template_path):
    if template_path and os.path.exists(template_path):
        try:
            return Presentation(template_path)
        except Exception:
            # 模板损坏，降级为空白创建
            return Presentation()
    return Presentation()
```

python-pptx 的占位符类型体系（`PP_PLACEHOLDER` 枚举）：

| PP_PLACEHOLDER 类型 | 值 | 对应 PageContent 字段 |
|---------------------|---|----------------------|
| TITLE | 1 | `title` |
| CENTER_TITLE | 3 | `title` |
| BODY | 2 | `bullets` |
| OBJECT | 7 | `bullets`（通用内容） |
| SUBTITLE | 4 | `subtitle` |
| CHART | 8 | `chart_data` |
| PICTURE | 18 | `image` |
| TABLE | 12 | `diagram_data`（表格类型） |

```python
def _fill_template_placeholder(slide, placeholder, content):
    """
    1. 识别模板占位符类型（通过 placeholder.placeholder_format.type）
    2. 按 PP_PLACEHOLDER 枚举映射到 PageContent 中对应字段
    3. 填充内容，保留模板的格式设置
    4. 模板占位符不够时，手动添加 shape 补充
    """
```

**LOGO 插入逻辑**：

> **实现约束**：python-pptx 默认模板的 slide_master 只有 placeholder shapes（title/text/date/footer/slide_number），没有图片 shape。企业模板的 LOGO 可能在：1) Slide Master 的非 placeholder 图片 shape 中；2) Slide Master 的 PICTURE 类型 placeholder 中；3) 某个 Slide Layout 中；4) 作为背景水印。检测规则（按优先级）：
>
> 1. 遍历 `slide_master.shapes`，找 `shape.shape_type == MSO_SHAPE_TYPE.PICTURE` 且名称含 "logo" 的 shape
> 2. 遍历 `slide_master.shapes`，找 `shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE` 的 shape
> 3. 遍历每个 `slide_layout.shapes`，重复上述检测（某些模板把 LOGO 放在 layout 而非 master）
> 4. 如果检测到模板已有 LOGO，记录其位置/尺寸到 BrandSpec.logo，EnterpriseRenderer 无需手动插入

```python
def _insert_logo(slide, logo_path, logo_spec):
    """
    1. 模板 master 上已有 LOGO（非 placeholder 图片 shape）→ 自动继承，无需操作
    2. 模板无 LOGO 但用户提供了 logo 文件 → 按 logo_spec 插入
    3. logo_spec.skip_slides 中的页面跳过
    """
```

**密度调整逻辑**：

```python
def _apply_density(text_frame, density_level):
    """
    density 1-3 (低): body=18pt, line_spacing=1.5x, bullets=2-3
    density 4-6 (中): body=14-16pt, line_spacing=1.2x, bullets=4-5
    density 7-10 (高): body=12-13pt, line_spacing=1.0x, bullets=6-8
    """
```

---

## 4. 密集布局系统

### 4.1 新增布局

面向 education/training 等高信息量场景，在 LayoutRegistry 中新增：

| Layout ID | Name | 适用 goal | 特点 |
|-----------|------|----------|------|
| 12 | Dense Text | concept, content | 标题 + 大面积正文区，6-8 bullet，小字号 |
| 13 | Table Layout | data, comparison | 标题 + 表格（可配置行列数） |
| 14 | Two Column Dense | comparison, before-after | 标题 + 左右两列各 4-5 bullet |
| 15 | Code Block | code-demo | 标题 + 代码区 + 输出/说明区 |
| 16 | Step Process | steps, procedure | 标题 + 3-5 步骤（编号+标题+描述） |
| 17 | Q&A Exercise | exercise, assessment | 题目区 + 提示区 + 答案区 |
| 18 | Outline / TOC | outline, agenda | 标题 + 章节列表 + 页码 |
| 19 | Key Takeaways | summary, takeaways | 标题 + 3-5 要点卡片 |
| 20 | Diagram Focus | process, pipeline, funnel, timeline, hierarchy, cycle, swot, comparison, pyramid, architecture | 标题 + 图形区 + 底部注释 |
| 21 | Diagram + Text | process-detail, architecture-detail | 标题 + 左图形区 + 右文字区 |
| 22 | Table Focus | data-table, specification | 标题 + 表格 |

### 4.2 密度参数实现

`density` 参数（1-10）影响渲染时的字号、间距、内容量：

| density | body字号 | 标题字号 | 行距 | bullet数 | 页边距 | 图片占比 |
|---------|---------|---------|------|---------|--------|---------|
| 1-3 | 18pt | 32pt | 1.5x | 2-3 | 宽松(1.0") | 40%+ |
| 4-6 | 14-16pt | 28pt | 1.2x | 4-5 | 适中(0.7") | 20-30% |
| 7-10 | 12-13pt | 24pt | 1.0x | 6-8 | 紧凑(0.5") | <10% |

**density 与 business_mode 的默认关系**：

| business_mode | 默认 density | 可调范围 |
|---------------|-------------|---------|
| pitch | 4 | 1-7 |
| education | 7 | 4-10 |
| training | 6 | 4-9 |
| report | 8 | 5-10 |

用户可通过 `--density` 覆盖默认值。

### 4.3 教学动画序列

教育场景的动画不是装饰，而是教学步骤：

| 动画类型 | 用途 | 实现方式 |
|---------|------|---------|
| Click-by-click reveal | 逐步展示要点 | 每个 bullet 独立 appear 动画 |
| Step-by-step build | 流程图逐步构建 | 每个步骤依次 appear |
| Code line highlight | 代码逐行高亮 | 每行独立 appear + 颜色变化 |
| Answer reveal | 答案延迟显示 | 答案区 appear on click |

> **实现约束**：python-pptx 没有元素级动画 API。当前 `_apply_transition` 仅实现幻灯片切换（slide-level），不是元素级动画（shape-level）。Click-by-click reveal 等教学动画需要在 `<p:timing>` XML 中为每个段落/shape 添加独立的 click effect，涉及大量 XML 构造。Phase G 中标注为研究性任务，需先验证 PPT 打开后的动画效果。

---

## 5. brand.json 规范

### 5.1 完整结构

```json
{
  "colors": {
    "primary": "#1A3C6E",
    "on-primary": "#FFFFFF",
    "secondary": "#64748B",
    "accent": "#2ECC71",
    "background": "#FFFFFF",
    "foreground": "#333333",
    "muted": "#F1F5F9",
    "muted-foreground": "#666666",
    "border": "#E2E8F0",
    "destructive": "#EF4444"
  },
  "fonts": {
    "heading": "Arial",
    "body": "Calibri"
  },
  "spacing": {
    "title_size_pt": 28,
    "body_size_pt": 14,
    "line_spacing": 1.2,
    "margins_inches": 0.8,
    "bullet_spacing_pt": 8
  },
  "logo": {
    "position": "top_right",
    "width_inches": 1.2,
    "on_every_slide": true,
    "skip_slides": ["title-slide", "cta-closing"]
  },
  "layout_mapping": {
    "hook": 0,
    "content": 2,
    "features": 3,
    "metrics": 4,
    "cta": 5
  },
  "density": 5,
  "animation_style": "conservative"
}
```

### 5.2 最小可用示例

只写想覆盖的字段，其余从模板提取或使用默认值。字段名与完整结构一致，未声明的字段不影响其他字段的提取逻辑：

```json
{
  "colors": {
    "primary": "#1A3C6E",
    "on-primary": "#FFFFFF"
  },
  "logo": {
    "position": "top_right"
  }
}
```

> **键名约定**：brand.json 的 `colors` 键名必须与 ThemeMapper 的颜色角色一致（`primary`/`on-primary`/`secondary`/`accent`/`background`/`foreground`/`muted`/`muted-foreground`/`border`/`destructive`），不使用 `text`/`text_secondary` 等非标准名。未声明的颜色角色从模板提取或使用默认值。

### 5.3 字段优先级

```
brand.json 字段 > template.pptx 提取 > 系统默认值
```

每个字段独立生效，未声明的字段不影响其他字段的提取逻辑。

---

## 6. content.json 规范

### 6.1 完整结构

```json
{
  "meta": {
    "title": "AI招聘平台产品介绍",
    "subtitle": "颠覆传统招聘方式",
    "author": "HR Tech Inc.",
    "date": "2026-07-11"
  },
  "slides": [
    {
      "goal": "hook",
      "title": "颠覆传统招聘方式",
      "subtitle": "AI驱动的下一代人才匹配平台",
      "image": "hero.png"
    },
    {
      "goal": "problem",
      "title": "传统招聘的三大痛点",
      "bullets": [
        "简历筛选耗时：HR平均每职位花费23小时",
        "匹配精度低：仅12%的面试转化为offer",
        "人才流失快：40%新员工6个月内离职"
      ]
    },
    {
      "goal": "features",
      "title": "核心功能",
      "cards": [
        {"title": "智能匹配", "body": "AI算法精准匹配", "image": "dashboard.png"},
        {"title": "自动筛选", "body": "简历自动评分排序"},
        {"title": "数据洞察", "body": "招聘全流程可视化"}
      ]
    }
  ]
}
```

### 6.2 图片引用

`image` 字段值为项目文件夹中的文件名（相对路径），系统自动解析为绝对路径。

---

## 7. CLI 参数设计

### 7.1 新增参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `--project` | str | 项目文件夹路径，触发 Enterprise Pipeline |
| `--business-mode` | choice | 业务模式：pitch/education/training/report |
| `--review` | flag | 启用确认环节（默认不启用，需显式指定） |
| `--review-file` | str | 方案输出到 JSON 文件（可编辑后重新运行，需配合 --review） |
| `--density` | int 1-10 | **扩展现有参数**：Enterprise Pipeline 中覆盖 business_mode 默认值 |
| `--output-version` | int | 指定输出版本号（覆盖已有版本，需配合 --project） |
| `--from-version` | int | 基于指定版本的 meta.json 上下文修订（需配合 --project） |
| `--pages` | str | 指定操作的页面范围（需配合 --from-version），见 §7.5 |
| `--history` | flag | 显示版本历史，不生成 PPT（需配合 --project，query 可选） |

### 7.2 使用示例

```bash
# 企业场景：基于模板+品牌规范（输出到 my_project/output/v1/）
python -m ppt_pro_max "AI招聘平台产品介绍" --project my_project/

# 修订：基于 v1 的上下文微调（输出到 my_project/output/v2/）
python -m ppt_pro_max "调整第5页布局" --project my_project/ --from-version 1

# 页面级修订：只重新生成第3页和第5页（其余页面保持原样）
python -m ppt_pro_max "更新竞品对比和增长数据" --project my_project/ --from-version 2 --pages 3,5

# 页面级修订：重新生成第7-9页（范围语法）
python -m ppt_pro_max "优化团队和CTA页" --project my_project/ --from-version 2 --pages 7-9

# 页面级修订：混合语法——第1页、第3-5页、第8页
python -m ppt_pro_max "调整多个页面" --project my_project/ --from-version 2 --pages 1,3-5,8

# 页面级新增：在第4页后插入新页（--pages 指定插入位置）
python -m ppt_pro_max "新增技术架构页" --project my_project/ --from-version 2 --pages +5

# 页面级删除：删除第6页
python -m ppt_pro_max "删除定价页" --project my_project/ --from-version 2 --pages -6

# 页面级移动：把第3页移到第5页后面
python -m ppt_pro_max "调整页面顺序" --project my_project/ --from-version 2 --pages 3>5

# 页面级交换：交换第3页和第7页的位置
python -m ppt_pro_max "调整页面顺序" --project my_project/ --from-version 2 --pages "3<>7"

# 教育场景：高密度+教学策略
python -m ppt_pro_max "Python基础教程" --project edu_project/ --business-mode education --density 8

# 企业内训：模板约束+培训策略
python -m ppt_pro_max "新员工入职培训" --project training_project/ --business-mode training

# 输出方案供确认
python -m ppt_pro_max "年度报告" --project report_project/ --review --review-file proposal.json

# 查看版本历史
python -m ppt_pro_max --project my_project/ --history

# 覆盖 v2（对 v2 不满意，重新生成）
python -m ppt_pro_max "AI招聘平台产品介绍" --project my_project/ --output-version 2

# FreeStyle（现有逻辑，零改动，输出到当前目录）
python -m ppt_pro_max "AI startup pitch" --style "dark cyberpunk" --fetch-images
```

### 7.3 向后兼容

所有现有参数保持不变。`--project` 不传时走 FreeStyle Pipeline，行为与当前完全一致。

### 7.4 Argparse 设计细节

**`query` 参数可选化**：`--history` 模式下不需要 query，但 FreeStyle 模式下 query 仍为必填。方案：

```python
# cli.py 修改
parser.add_argument('query', nargs='?', default=None,
                    help='PPT 主题描述（--history 模式下可省略）')

# 参数验证逻辑（在 parse_args 之后）
if not args.history and args.query is None:
    parser.error('query is required unless --history is specified')
```

**`--output-version` 与 `--from-version` 互斥**：两者不能同时指定（一个覆盖输出，一个基于旧版修订）：

```python
group = parser.add_mutually_exclusive_group()
group.add_argument('--output-version', type=int, metavar='N',
                   help='指定输出版本号（覆盖已有版本）')
group.add_argument('--from-version', type=int, metavar='N',
                   help='基于指定版本的 meta.json 上下文修订')
```

**`--review-file` 依赖 `--review`**：

```python
if args.review_file and not args.review:
    parser.error('--review-file requires --review')
```

**`--pages` 依赖 `--from-version`**：

```python
if args.pages and not args.from_version and not args.history:
    parser.error('--pages requires --from-version (or --history for query mode)')
```

### 7.5 `--pages` 页面范围语法

`--pages` 支持灵活的页面范围语法，用于指定页面级增删改移换操作（需配合 `--from-version`，或 `--history` 查看详情）：

**语法规则**：

| 语法 | 含义 | 操作类型 |
|------|------|---------|
| `3` | 第3页 | 重新生成（改） |
| `3,5,8` | 第3、5、8页 | 重新生成（改） |
| `3-5` | 第3到第5页 | 重新生成（改） |
| `1,3-5,8` | 第1、3-5、8页 | 重新生成（改） |
| `+5` | 在第5页**后**插入新页 | 新增（增） |
| `+5,+8` | 在第5页后、第8页后各插入新页 | 新增（增） |
| `-6` | 删除第6页 | 删除（删） |
| `-3,-7` | 删除第3页和第7页 | 删除（删） |
| `3>5` | 第3页移到第5页后面 | 移动（移） |
| `3<>7` | 交换第3页和第7页 | 交换（换） |
| `3-5,+6,-8` | 重新生成3-5页，第6页后插入，删除第8页 | 混合操作 |

> **页码约定**：所有页码均为 **1-based**，始终指**原始文档**中的页码。即使操作中包含删除/移动，后续操作的页码仍基于原始文档。例如 `--pages "-3,5"` 表示"删除原始第3页，修改原始第5页"——序列变换算法（§9.6.4）负责正确处理索引映射。

**解析实现**：

```python
from dataclasses import dataclass

@dataclass
class PageOp:
    """单页操作指令（统一数据结构，§7.5 定义，§9.6.5/§9.6.7 引用）"""
    page: int               # 1-based 页码，指原始文档
    action: str             # "modify" | "insert" | "delete" | "move" | "swap"
    target: int = None      # 仅 move/swap: 目标页码（1-based）
    insert_after: bool = True  # 仅 insert: True=在page后插入

def parse_pages(pages_str: str, num_slides: int = 0) -> list[PageOp]:
    """
    解析 --pages 参数字符串为操作列表
    
    Args:
        pages_str: --pages 参数值
        num_slides: 原始文档页数，用于验证页码范围（0=不验证）
    
    示例:
      "3,5"       → [PageOp(3,modify), PageOp(5,modify)]
      "3-5"       → [PageOp(3,modify), PageOp(4,modify), PageOp(5,modify)]
      "+5"        → [PageOp(5,insert)]
      "-6"        → [PageOp(6,delete)]
      "3>5"       → [PageOp(3,move,target=5)]
      "3<>7"      → [PageOp(3,swap,target=7)]
      "1,3-5,+6,-8,3>5,3<>7" → 混合操作
    """
    ops = []
    for token in pages_str.split(","):
        token = token.strip()
        if not token:
            continue
        
        # 交换: N<>M
        if "<>" in token:
            parts = token.split("<>")
            a, b = int(parts[0]), int(parts[1])
            ops.append(PageOp(page=a, action="swap", target=b))
        
        # 移动: N>M
        elif ">" in token and not token.startswith(">"):
            parts = token.split(">")
            a, b = int(parts[0]), int(parts[1])
            ops.append(PageOp(page=a, action="move", target=b))
        
        # 新增: +N
        elif token.startswith("+"):
            page = int(token[1:])
            ops.append(PageOp(page=page, action="insert", insert_after=True))
        
        # 删除: -N
        elif token.startswith("-") and token[1:].isdigit():
            page = int(token[1:])
            ops.append(PageOp(page=page, action="delete"))
        
        # 范围: N-M
        elif "-" in token and not token.startswith("-"):
            parts = token.split("-")
            start, end = int(parts[0]), int(parts[1])
            for p in range(start, end + 1):
                ops.append(PageOp(page=p, action="modify"))
        
        # 单页: N
        elif token.isdigit():
            ops.append(PageOp(page=int(token), action="modify"))
    
    # 页码范围验证
    if num_slides > 0:
        for op in ops:
            if op.action in ("modify", "delete", "move", "swap"):
                if op.page < 1 or op.page > num_slides:
                    raise ValueError(f"页码 {op.page} 超出范围（1-{num_slides}）")
            if op.action in ("move", "swap") and op.target is not None:
                if op.target < 1 or op.target > num_slides:
                    raise ValueError(f"目标页码 {op.target} 超出范围（1-{num_slides}）")
            if op.action == "insert":
                if op.page < 1 or op.page > num_slides:
                    raise ValueError(f"插入位置 {op.page} 超出范围（1-{num_slides}）")
    
    return ops
```

> **注意**：`parse_pages()` 返回的操作列表不排序——排序由 `compute_target_sequence()`（§9.6.4）按操作类型决定执行顺序（SWAP → MOVE → DELETE → INSERT）。

---

## 8. 降级策略

Enterprise Pipeline 内部根据资产完整度自动降级：

| 有模板 | 有brand | 有content | 有图片 | 工作模式 |
|-------|---------|----------|--------|---------|
| ✓ | ✓ | ✓ | ✓ | 全约束模式：模板渲染+品牌规范+真实内容+图片匹配 |
| ✓ | ✗ | ✗ | ✗ | 模板模式：从模板提取 brand，自动生成内容 |
| ✗ | ✓ | ✗ | ✓ | 品牌约束模式：空白创建但遵循 brand 颜色/字体 |
| ✗ | ✗ | ✓ | ✗ | 内容模式：空白创建，使用用户内容 |
| ✗ | ✗ | ✗ | ✗ | 最小模式：空白创建 + Review Gate（等价于 FreeStyle + 确认环节） |

**原则：任何组合都不报错，只是精度不同。资产越全，输出越贴合规范；资产越少，自由度越高。**

---

## 9. 输出目录与版本管理

### 9.1 固定输出位置

Enterprise Pipeline 的输出固定在项目目录下的 `output/` 子目录，不再随机生成路径。

**规则**：

| 场景 | 输出路径 |
|------|---------|
| Enterprise（`--project my_project/`） | `my_project/output/v{N}/presentation.pptx` |
| Enterprise + `--output custom.pptx` | `custom.pptx`（用户显式指定，优先级最高） |
| FreeStyle（无 `--project`） | 当前目录 + 时间戳（现有逻辑不变） |

**自动创建**：`output/` 和 `v{N}/` 在生成时自动创建，无需用户手动建目录。

### 9.2 版本编号规则

每次生成自动递增版本号，从 v1 开始：

```
my_project/output/
├── v1/                    # 第1次生成
│   ├── presentation.pptx
│   └── meta.json
├── v2/                    # 第2次生成（修订）
│   ├── presentation.pptx
│   └── meta.json
├── v3/                    # 第3次生成
│   ├── presentation.pptx
│   └── meta.json
└── latest -> v3/          # 始终指向最新版
```

**版本号确定逻辑**：

```python
def next_version(output_dir: str) -> int:
    """扫描 output/ 下已有的 v{N}/ 目录，返回 max(existing) + 1

    注意：
    - latest junction 以 "l" 开头，不会被 "v" + digit 匹配，无需额外排除。
    - 使用 max+1 而非顺序填充：如果 v1,v2,v7 存在，返回 v8（不是 v3）。
      这保证永远不会覆盖已删除的版本号，即使中间有间隔。
    """
    existing = []
    for entry in os.listdir(output_dir):
        if entry.startswith("v") and entry[1:].isdigit():
            existing.append(int(entry[1:]))
    return max(existing, default=0) + 1
```

**用户指定版本号**：

```bash
# 自动递增（默认）
python -m ppt_pro_max "query" --project my_project/

# 指定版本号（覆盖已有版本）
python -m ppt_pro_max "query" --project my_project/ --output-version 2

# 基于指定版本的内容修订（读取 v2 的 meta.json 恢复上下文）
python -m ppt_pro_max "query" --project my_project/ --from-version 2
```

### 9.3 版本元数据 meta.json

每个版本目录下自动生成 `meta.json`，记录本次生成的完整上下文，支持回溯和复现：

```json
{
  "version": 2,
  "created_at": "2026-07-11T14:30:00",
  "query": "AI招聘平台产品介绍",
  "business_mode": "pitch",
  "density": 5,
  "strategy": "YC Seed Deck",
  "page_count": 10,
  "assets_used": {
    "template": "template.pptx",
    "logo": "logo.png",
    "brand_json": true,
    "content_json": true,
    "image_pool": ["hero.png", "dashboard.png", "team.jpg"]
  },
  "brand_spec": {
    "colors": {"primary": "#1A3C6E", "secondary": "#E8491D"},
    "fonts": {"heading": "Arial", "body": "Calibri"}
  },
  "story_plan": {
    "strategy": "YC Seed Deck",
    "pages": [
      {"position": 1, "goal": "hook", "emotion": "curiosity"},
      {"position": 2, "goal": "problem", "emotion": "frustration"}
    ]
  },
  "page_designs": [
    {"position": 1, "goal": "hook", "layout": "title-slide", "title": "颠覆传统招聘方式"},
    {"position": 2, "goal": "problem", "layout": "content-with-title", "title": "传统招聘的三大痛点"}
  ],
  "slides": [
    {
      "position": 1,
      "goal": "hook",
      "title": "颠覆传统招聘方式",
      "subtitle": "AI驱动的下一代人才匹配平台",
      "layout_id": 0,
      "layout_name": "Title Slide",
      "image": "hero.png",
      "bullets": null,
      "cards": null,
      "diagram_type": null,
      "diagram_data": null
    },
    {
      "position": 2,
      "goal": "problem",
      "title": "传统招聘的三大痛点",
      "subtitle": null,
      "layout_id": 2,
      "layout_name": "Content",
      "image": null,
      "bullets": ["简历筛选耗时", "匹配精度低", "人才流失快"],
      "cards": null,
      "diagram_type": null,
      "diagram_data": null
    }
  ],
  "review_gate": {
    "reviewed": true,
    "user_modifications": {
      "page_5_layout": "three-column-cards",
      "logo_skip_slides": ["title-slide"]
    }
  },
  "qa_result": {
    "passed": true,
    "checks": {"slide_count": true, "title_presence": true, "content_density": true}
  }
}
```

> **`slides[]` 字段说明**：v1.5 新增，记录每页的完整内容快照。这是页面级修订的关键——`--pages` 修订时，未指定的页面直接从 `slides[]` 恢复内容，无需重新生成。`slides[]` 的结构与 `content.json` 的 `slides[]` 一致，确保双向兼容。

**meta.json 的用途**：

| 用途 | 说明 |
|------|------|
| 回滚 | 用户说"还是 v1 好"，直接用 `output/v1/presentation.pptx` |
| 复现 | 读取 meta.json 还原完整生成上下文，基于旧版本微调 |
| 对比 | 对比 v1 和 v2 的 story_plan / page_designs 差异 |
| 审计 | 记录每次生成的参数和结果，便于追溯 |
| 修订接力 | `--from-version 2` 读取 v2 的 meta.json，在其基础上修改 |

### 9.4 latest 链接

`output/latest` 始终指向最新版本，方便快速访问：

- **Linux/Mac**：符号链接 `os.symlink()`
- **Windows**：目录 junction `subprocess.run(['mklink', '/J', target, source], shell=True, check=True)`，无需管理员权限
- **兼容降级**：如果创建链接失败，改为复制最新版 .pptx 到 `output/latest/presentation.pptx`

### 9.5 版本对比

```bash
# 查看版本历史
python -m ppt_pro_max --project my_project/ --history

# 输出:
# v1 | 2026-07-11 10:00 | 10页 | pitch | YC Seed Deck
# v2 | 2026-07-11 14:30 | 12页 | pitch | YC Seed Deck (修订: +2页, 改了第5页布局)
# v3 | 2026-07-12 09:15 | 12页 | pitch | YC Seed Deck (修订: 改了品牌色)
```

### 9.6 页面级修订（增删查改移换）

#### 9.6.1 设计目标

企业场景中，用户经常只需要调整某一页或某几页，而非全量重新生成。页面级修订的核心诉求：

| 诉求 | 说明 | 效率提升 |
|------|------|---------|
| **只改需要的页** | 10页中只改第5页，其余9页原样保留 | API 调用减少 90%，时间减少 90% |
| **风格统一** | 新生成/修改的页面必须与已有页面风格一致 | BrandSpec + meta.json 保证 |
| **插入新页** | 在现有 PPT 中插入新页面 | 无需全量重做 |
| **删除页面** | 去掉不需要的页面 | 无需全量重做 |
| **移动页面** | 把第3页移到第5页后面 | 无需全量重做 |
| **交换页面** | 交换第3页和第7页的位置 | 无需全量重做 |
| **成本可控** | 只对目标页面调用 LLM/图片生成 | Token 和 API 费用按比例降低 |

#### 9.6.2 操作类型

| 操作 | CLI 语法 | 说明 | 风格保证 |
|------|---------|------|---------|
| **改**（重新生成） | `--pages 3,5` | 重新生成指定页面的内容 | BrandSpec + 同 layout |
| **增**（插入新页） | `--pages +5` | 在第5页后插入新页面 | BrandSpec + EnterpriseDesignDecider |
| **删**（删除页面） | `--pages -6` | 删除第6页 | 无需渲染 |
| **移**（移动页面） | `--pages 3>5` | 把第3页移到第5页后面 | 原样复制，零风险 |
| **换**（交换页面） | `--pages 3<>7` | 交换第3页和第7页的位置 | 原样复制，零风险 |
| **查**（查看详情） | `--history --pages 3` | 输出第3页的 meta.json 详情 | 无需渲染 |

#### 9.6.3 核心算法：两阶段重建（Two-Pass Rebuild）

> **关键发现**：python-pptx 在删除 slide 后调用 `add_slide()`，新 slide 会复用已删除的 partname（如 `slide9.xml` 出现两次），导致 ZIP 冲突，保存后数据损坏。**不能在同一个 Presentation 对象上先删后增**。
>
> 解决方案：**两阶段重建**——Pass 1 修改源文件内容，Pass 2 从模板创建新 Presentation 并按目标顺序复制 slide。新 Presentation 的 partname 由 python-pptx 自动顺序分配，彻底避免冲突。

**算法流程**：

```
--from-version N --pages "3,5,+6,-8,3>5,3<>7"
         │
         ▼
  ┌─────────────────────────────────────────────────┐
  │  Pass 1: 内容修改（在源 Presentation 上操作）      │
  │  源文件: v{N}/presentation.pptx                   │
  │                                                   │
  │  1. 读取 v{N}/meta.json → 恢复上下文              │
  │  2. 打开 v{N}/presentation.pptx                   │
  │  3. 对 MODIFY 操作的页面：                         │
  │     - 修改 slide 中的文字/内容（in-place）          │
  │     - 保持 layout 不变                             │
  │  4. 保存为临时文件 temp_modified.pptx              │
  └─────────────────────────────────────────────────┘
         │
         ▼
  ┌─────────────────────────────────────────────────┐
  │  Pass 2: 结构重建（从模板创建新 Presentation）      │
  │                                                   │
  │  1. 计算目标 slide 序列（见 §9.6.4 序列变换算法）   │
  │  2. 打开 template.pptx → 删除所有默认 slide        │
  │     → 得到只有 master/layouts 的空壳              │
  │  3. 按目标序列逐个复制 slide：                     │
  │     - COPY: deepcopy 源 slide 的 shapes XML       │
  │       → 使用同名 layout（保证模板布局一致）         │
  │     - NEW: add_slide + EnterpriseRenderer 渲染    │
  │  4. 保存为 v{N+1}/presentation.pptx              │
  │  5. 更新 meta.json                                │
  └─────────────────────────────────────────────────┘
```

**为什么两阶段而非原地操作**：

| 方案 | 删除+插入 | 移动/交换 | 模板布局 | partname 冲突 |
|------|---------|---------|---------|-------------|
| 原地操作（delete+add+reorder） | ❌ partname 冲突 | ⚠️ sldIdLst 操作复杂 | ✅ 天然保留 | ❌ 已验证会损坏 |
| **两阶段重建** | ✅ 新 prs 自动编号 | ✅ 按序列复制 | ✅ 同名 layout 匹配 | ✅ 无冲突 |

#### 9.6.4 序列变换算法

所有操作最终归结为一个问题：**给定原始序列 [0,1,2,...,N-1]，计算目标序列**。

```python
def compute_target_sequence(num_slides: int, ops: list[PageOp]) -> tuple[list, dict]:
    """
    根据操作列表计算目标 slide 序列。
    
    所有 page 编号为 1-based，指原始文档中的页码。
    
    返回:
        target_order: 目标 slide 序列（0-based 源索引列表）
        new_slides: {插入位置: 内容} 字典
    
    算法步骤:
    1. 从原始序列 [0, 1, 2, ..., N-1] 开始
    2. 应用 MOVE: 将源索引从原位置移到目标位置
    3. 应用 SWAP: 交换两个源索引的位置
    4. 应用 DELETE: 移除指定源索引
    5. 应用 INSERT: 在指定位置标记新 slide
    """
    # 原始序列：每个元素是源 slide 的 0-based 索引
    sequence = list(range(num_slides))
    new_slides = {}
    
    # 操作执行顺序：SWAP → MOVE → DELETE → INSERT
    # 前三种只改变顺序/成员，不引入新元素，在同一个序列上操作
    # INSERT 引入新元素，最后处理
    
    # --- SWAP ---
    swap_ops = [o for o in ops if o.action == "swap"]
    for op in swap_ops:
        # op.page 和 op.target 是 1-based
        idx_a = op.page - 1
        idx_b = op.target - 1
        # 在 sequence 中找到这两个源索引的位置
        pos_a = sequence.index(idx_a)
        pos_b = sequence.index(idx_b)
        sequence[pos_a], sequence[pos_b] = sequence[pos_b], sequence[pos_a]
    
    # --- MOVE ---
    move_ops = [o for o in ops if o.action == "move"]
    for op in move_ops:
        # op.page 移到 op.target 后面（1-based）
        idx_source = op.page - 1
        idx_target = op.target - 1
        # 在 sequence 中找到源的位置
        pos_source = sequence.index(idx_source)
        # 移除源
        sequence.pop(pos_source)
        # 找到目标的新位置（因为移除源后位置可能变化）
        pos_target = sequence.index(idx_target)
        # 插入到目标后面
        sequence.insert(pos_target + 1, idx_source)
    
    # --- DELETE ---
    delete_ops = sorted([o for o in ops if o.action == "delete"], key=lambda o: o.page)
    for op in delete_ops:
        idx = op.page - 1
        if idx in sequence:
            sequence.remove(idx)
    
    # --- INSERT ---
    insert_ops = sorted([o for o in ops if o.action == "insert"], key=lambda o: o.page)
    for op in insert_ops:
        # op.page: 在第 op.page 页后插入（1-based，指原始文档页码）
        # 需要映射到当前 sequence 中的位置
        # 找到原始第 op.page 页在 sequence 中的位置
        idx_original = op.page - 1
        if idx_original in sequence:
            pos = sequence.index(idx_original)
            insert_pos = pos + 1
        else:
            # 原始页已被删除，插入到最近的未被删除的前一页后
            insert_pos = len(sequence)  # 末尾
        new_slides[insert_pos] = op  # 标记插入位置
    
    return sequence, new_slides
```

**场景验证**：

| 场景 | 原始序列 | 操作 | 目标序列 | 验证 |
|------|---------|------|---------|------|
| 删除第3、7页 | [0,1,2,3,4,5,6,7,8,9] | delete 3,7 | [0,1,3,4,5,7,8,9] | ✅ |
| 第3页移到第5页后 | [0,1,2,3,4,5,6,7] | move 3→5后 | [0,1,3,4,2,5,6,7] | ✅ |
| 交换第3、7页 | [0,1,2,3,4,5,6,7] | swap 3↔7 | [0,1,6,3,4,5,2,7] | ✅ |
| 删3,7 + 改2,5 + 第4页后插入 | [0,1,2,3,4,5,6,7,8,9] | delete 3,7; insert after 4 | [0,1,3,4,NEW,5,7,8,9] | ✅ |
| 第8页移到第1页后 + 删3 | [0,1,2,3,4,5,6,7,8,9] | move 8→1后; delete 3 | [0,7,1,3,4,5,8,9] | ✅ |

#### 9.6.5 `--pages` 语法

`--pages` 的完整语法定义和 `PageOp`/`parse_pages()` 实现见 §7.5。此处补充与序列变换算法的关系：

- `parse_pages()` 返回 `list[PageOp]`，不排序
- `compute_target_sequence()`（§9.6.4）按操作类型决定执行顺序：SWAP → MOVE → DELETE → INSERT
- 所有页码均为 1-based，指原始文档，不存在索引偏移问题

#### 9.6.6 风格一致性保证

页面级修订最大的风险是**新生成的页面与已有页面风格不一致**。保证机制：

```
┌─────────────────────────────────────────────────────┐
│              风格一致性三层保证                        │
├─────────────────────────────────────────────────────┤
│                                                     │
│  Layer 1: BrandSpec（颜色/字体/间距）                 │
│  ─────────────────────────────────────               │
│  • 从 meta.json.brand_spec 恢复                      │
│  • 新页面使用相同的 BrandSpec 渲染                     │
│  • 保证：主色、辅色、字号、字体名完全一致               │
│                                                     │
│  Layer 2: Layout 映射（布局结构）                      │
│  ─────────────────────────────────────               │
│  • COPY 操作：按源 slide 的 layout name 匹配          │
│    新 prs 中的同名 layout，保证模板布局一致             │
│  • MODIFY 操作：默认保持原 layout                     │
│  • INSERT 操作：由 EnterpriseDesignDecider 选择       │
│    与相邻页面风格协调的 layout                         │
│  • 保证：页面结构（标题位置、内容区域）一致             │
│                                                     │
│  Layer 3: 渲染参数（密度/装饰/动画）                   │
│  ─────────────────────────────────────               │
│  • 从 meta.json 恢复 density、business_mode           │
│  • 新页面使用相同的渲染参数                            │
│  • 保证：字号层级、间距、装饰风格一致                   │
│                                                     │
└─────────────────────────────────────────────────────┘
```

**Layout 匹配实现**（Pass 2 中复制 slide 时）：

```python
def _match_layout(new_prs: Presentation, source_slide) -> SlideLayout:
    """在目标 Presentation 中找到与源 slide 同名的 layout"""
    source_layout_name = source_slide.slide_layout.name
    for layout in new_prs.slide_layouts:
        if layout.name == source_layout_name:
            return layout
    # 找不到同名 layout → 降级为空白布局
    return new_prs.slide_layouts[6]
```

#### 9.6.7 具体实现

```python
class PageRevisionEngine:
    """页面级修订引擎：两阶段重建，保证风格一致性"""

    def revise(self, version_path: str, template_path: str, 
               meta: dict, ops: list[PageOp],
               brand_spec: BrandSpec, query: str, **kwargs) -> str:
        """
        执行页面级修订，返回新版本路径。
        
        两阶段重建：
        Pass 1: 修改源文件内容（in-place）
        Pass 2: 从模板重建，按目标序列复制 slide
        """
        slides_meta = meta.get("slides", [])
        density = meta.get("density", 5)
        business_mode = meta.get("business_mode", "pitch")
        num_slides = len(slides_meta)

        # ─── Pass 1: 内容修改 ───
        source = Presentation(version_path)
        
        modify_ops = [o for o in ops if o.action == "modify"]
        for op in modify_ops:
            idx = op.page - 1
            old_meta = slides_meta[idx]
            new_content = self._generate_modified_content(
                query, old_meta, brand_spec, density, business_mode, **kwargs
            )
            # 修改 slide 内容（in-place，不改结构）
            self._update_slide_content(source.slides[idx], new_content, brand_spec, density)
            slides_meta[idx] = {**old_meta, **new_content.to_meta()}
        
        # 保存修改后的源文件
        temp_path = version_path + ".temp"
        source.save(temp_path)

        # ─── Pass 2: 结构重建 ───
        # 计算目标序列
        target_order, new_slide_ops = compute_target_sequence(num_slides, ops)
        
        # 从模板创建新 Presentation
        if template_path and os.path.exists(template_path):
            new_prs = Presentation(template_path)
            # 删除模板自带的所有 slide（保留 master/layouts）
            # 安全：这是从模板新建的空壳，删除模板自带 slide 后再 add_slide
            # 不会产生 partname 冲突（与源文件上的先删后增不同）
            while len(new_prs.slides) > 0:
                rId = new_prs.slides._sldIdLst[0].get(qn('r:id'))
                new_prs.part.drop_rel(rId)
                new_prs.slides._sldIdLst.remove(new_prs.slides._sldIdLst[0])
        else:
            new_prs = Presentation()
            new_prs.slide_width = source.slide_width
            new_prs.slide_height = source.slide_height
        
        # 重新打开修改后的源文件（只读）
        modified_source = Presentation(temp_path)
        
        # 按目标序列构建新 Presentation
        # 注意：LOGO 位置从 BrandSpec.logo 恢复，新插入的页面按 logo_spec
        # 决定是否插入 LOGO（skip_slides 中的页面跳过）
        insert_count = 0
        for pos, source_idx in enumerate(target_order):
            # 检查是否需要在此位置插入新 slide
            while new_slide_ops and min(new_slide_ops.keys()) <= pos + insert_count:
                insert_pos = min(new_slide_ops.keys())
                op = new_slide_ops.pop(insert_pos)
                # 生成并渲染新页面
                adjacent_meta = slides_meta[min(source_idx, len(slides_meta) - 1)]
                new_content = self._generate_new_page(
                    query, adjacent_meta, brand_spec, density, business_mode, **kwargs
                )
                layout = self._select_layout(new_prs, adjacent_meta, brand_spec)
                new_slide = new_prs.slides.add_slide(layout)
                self._render_to_slide(new_slide, new_content, brand_spec, density)
                insert_count += 1
            
            # 复制源 slide
            source_slide = modified_source.slides[source_idx]
            target_layout = self._match_layout(new_prs, source_slide)
            new_slide = new_prs.slides.add_slide(target_layout)
            # 清除默认内容
            for shape in list(new_slide.shapes):
                shape._element.getparent().remove(shape._element)
            # 深拷贝所有 shapes
            for shape in source_slide.shapes:
                el = copy.deepcopy(shape._element)
                new_slide.shapes._spTree.append(el)
        
        # 处理末尾的插入
        while new_slide_ops:
            insert_pos = min(new_slide_ops.keys())
            op = new_slide_ops.pop(insert_pos)
            adjacent_meta = slides_meta[-1]
            new_content = self._generate_new_page(
                query, adjacent_meta, brand_spec, density, business_mode, **kwargs
            )
            layout = self._select_layout(new_prs, adjacent_meta, brand_spec)
            new_slide = new_prs.slides.add_slide(layout)
            self._render_to_slide(new_slide, new_content, brand_spec, density)
        
        # 清理临时文件
        os.remove(temp_path)
        
        return new_prs

    def _update_slide_content(self, slide, content, brand_spec, density):
        """修改 slide 内容（in-place，不改结构）"""
        # 修改标题
        if slide.shapes.title and content.title:
            slide.shapes.title.text = content.title
        # 修改正文
        for shape in slide.shapes:
            if shape.has_text_frame and shape != slide.shapes.title:
                if content.bullets:
                    for i, bullet in enumerate(content.bullets):
                        if i < len(shape.text_frame.paragraphs):
                            shape.text_frame.paragraphs[i].text = bullet
        # 修改图片（如有）
        if content.image:
            # 替换图片 shape 的内容
            pass  # 具体实现见 EnterpriseRenderer

    def _match_layout(self, new_prs, source_slide):
        """在目标 prs 中匹配同名 layout"""
        source_layout_name = source_slide.slide_layout.name
        for layout in new_prs.slide_layouts:
            if layout.name == source_layout_name:
                return layout
        # 降级为最后一个布局（通常是 Blank，比硬编码 [6] 更安全）
        return new_prs.slide_layouts[-1]

    def _select_layout(self, new_prs, adjacent_meta, brand_spec):
        """为新插入的页面选择 layout"""
        # 由 EnterpriseDesignDecider 根据相邻页面和 brand_spec 决定
        layout_mapping = brand_spec.layout_mapping if brand_spec and brand_spec.layout_mapping else {}
        goal = adjacent_meta.get("goal", "content")
        layout_idx = layout_mapping.get(goal)
        if layout_idx is not None and layout_idx < len(new_prs.slide_layouts):
            return new_prs.slide_layouts[layout_idx]
        return new_prs.slide_layouts[6]

    def _generate_modified_content(self, query, old_meta, brand_spec, 
                                   density, business_mode, **kwargs):
        """基于修改指令和旧页面内容生成新内容"""
        context = {
            "original_goal": old_meta.get("goal"),
            "original_title": old_meta.get("title"),
            "original_bullets": old_meta.get("bullets"),
            "modification_instruction": query,
            "brand_spec": brand_spec,
            "density": density,
            "business_mode": business_mode,
        }
        return ContentGenerator.generate_page(context, **kwargs)

    def _generate_new_page(self, query, adjacent_meta, brand_spec,
                           density, business_mode, **kwargs):
        """生成新插入页面的内容"""
        context = {
            "new_page_query": query,
            "previous_goal": adjacent_meta.get("goal"),
            "previous_title": adjacent_meta.get("title"),
            "brand_spec": brand_spec,
            "density": density,
            "business_mode": business_mode,
        }
        return ContentGenerator.generate_page(context, **kwargs)
```

#### 9.6.8 场景分析

**场景1：删除后索引偏移**

```
原始: [P1, P2, P3, P4, P5, P6, P7, P8]
操作: --pages -3,-5

序列变换:
  delete 3 → [0,1,3,4,5,6,7]  (移除源索引2)
  delete 5 → [0,1,3,4,6,7]    (移除源索引4)
  
结果: P1, P2, P4, P6, P7, P8 ✅

关键：delete 操作移除的是"源索引"而非"当前位置"，
所以不存在索引偏移问题。compute_target_sequence 
在原始序列上操作，每个 page 编号始终指原始文档。
```

**场景2：移动 + 删除**

```
原始: [P1, P2, P3, P4, P5, P6, P7, P8]
操作: --pages 3>5,-7

序列变换:
  move 3→5后: [0,1,3,4,2,5,6,7]  (源索引2移到源索引4后面)
  delete 7:   [0,1,3,4,2,5,7]     (移除源索引6)
  
结果: P1, P2, P4, P5, P3, P6, P8 ✅
```

**场景3：交换 + 修改 + 插入**

```
原始: [P1, P2, P3, P4, P5, P6, P7, P8]
操作: --pages 3<>7,2,+4

序列变换:
  swap 3↔7: [0,1,6,3,4,5,2,7]  (源索引2和6互换位置)
  insert after 4: 在源索引3后面插入 NEW
  
Pass 1: 修改源索引1（P2）的内容
Pass 2: 按序列 [0,1,6,3,NEW,4,5,2,7] 复制
  
结果: P1, MOD_P2, P7, P4, NEW, P5, P6, P3, P8 ✅
```

**场景4：多次移动（链式移动）**

```
原始: [P1, P2, P3, P4, P5, P6]
操作: --pages 2>4,3>6

序列变换:
  move 2→4后: [0,2,3,1,4,5]  (源索引1移到源索引3后面)
  move 3→6后: [0,2,1,4,3,5]  (源索引2移到源索引5后面)
  
结果: P1, P3, P2, P5, P4, P6 ✅

注意：move 操作按出现顺序依次执行，
每次 move 基于上一次 move 的结果。
```

**场景5：全量重生成（无 --pages）**

```
操作: --from-version 2（不带 --pages）

等价于: 所有页面都是 MODIFY
target_order = [0,1,2,...,N-1]（原序）
new_slides = {}（无插入）

Pass 1: 所有页面重新生成内容
Pass 2: 按原序复制到新 prs

结果: 全量重生成，但使用两阶段重建保证无 partname 冲突 ✅
```

#### 9.6.9 成本对比

| 操作 | LLM 调用 | 图片生成 | 时间 | Token 消耗 |
|------|---------|---------|------|-----------|
| 全量重生成（10页） | 10次 | 3-5次 | ~60s | ~10K |
| 改2页（`--pages 3,5`） | 2次 | 0-1次 | ~12s | ~2K |
| 插入1页（`--pages +5`） | 1次 | 0-1次 | ~6s | ~1K |
| 删除1页（`--pages -6`） | 0次 | 0次 | <1s | 0 |
| 移动1页（`--pages 3>5`） | 0次 | 0次 | <1s | 0 |
| 交换2页（`--pages 3<>7`） | 0次 | 0次 | <1s | 0 |

#### 9.6.10 修订记录

每次页面级修订在 meta.json 中记录变更详情（修订版本的 meta.json 包含 §9.3 的所有字段，以下仅展示新增/变更字段）：

```json
{
  "version": 3,
  "parent_version": 2,
  "revision_type": "page-level",
  "revision_ops": [
    {"page": 3, "action": "modify", "reason": "更新竞品对比数据"},
    {"page": 5, "action": "modify", "reason": "换用三栏布局"},
    {"page": 6, "action": "insert", "reason": "新增技术架构页"},
    {"page": 8, "action": "delete", "reason": "删除定价页"},
    {"page": 3, "action": "move", "target": 5, "reason": "第3页移到第5页后"},
    {"page": 3, "action": "swap", "target": 7, "reason": "交换第3页和第7页"}
  ],
  "unchanged_pages": [1, 2, 4, 7, 9, 10]
}
```

`--history` 输出增强：

```
v1 | 2026-07-11 10:00 | 10页 | pitch | YC Seed Deck
v2 | 2026-07-11 14:30 | 10页 | pitch | YC Seed Deck | 全量修订
v3 | 2026-07-11 16:00 | 9页  | pitch | YC Seed Deck | 页面级: 改3,5 +增6 -删8 移3>5 换3<>7
```

### 9.7 CLI 新增参数

| 参数 | 类型 | 说明 |
|------|------|------|
| `--output-version` | int | 指定输出版本号（覆盖已有版本） |
| `--from-version` | int | 基于指定版本的 meta.json 上下文修订 |
| `--pages` | str | 页面级操作范围（需配合 --from-version 或 --history） |
| `--history` | flag | 显示版本历史，不生成 PPT |

> 与 §7.1 中定义一致，此处不再重复描述。

### 9.8 与 FreeStyle 的关系

版本管理仅对 Enterprise Pipeline 生效。FreeStyle 保持现有行为（时间戳文件名，无版本管理），零改动。

---

## 10. 目录结构

```
src/ppt_pro_max/
├── __init__.py                    # 入口，根据 --project 分流
├── cli.py                         # 新增 --project/--business-mode/--review 等参数
│
├── planner/
│   └── story_planner.py           # 扩展：新增 education/training 策略
│
├── decider/
│   └── design_decider.py          # 不动，FreeStyle 专用
│
├── content/
│   └── content_generator.py       # 扩展：图片池匹配逻辑
│
├── renderer/
│   ├── ppt_renderer.py            # 不动，FreeStyle 专用
│   ├── theme_composer.py          # 不动，共享
│   ├── layout_registry.py         # 扩展：新增密集布局（ID 12-19）+ 图形/表格布局（ID 20-22）
│   ├── image_fetcher.py           # 不动，共享
│   ├── chart_builder.py           # 不动，共享
│   ├── effects.py                 # 不动，共享
│   ├── theme_mapper.py            # 扩展：BrandSpec 映射能力
│   ├── diagram_engine.py          # ★ 新增：图形引擎主入口
│   └── diagram/                   # ★ 新增：图形类型实现
│       ├── __init__.py
│       ├── base.py                # BaseDiagram 基类
│       ├── layout_engine.py       # 坐标计算、自适应排版
│       ├── text_measurer.py       # 文字尺寸估算
│       ├── data_splitter.py       # 分页拆分
│       ├── diagram_style.py       # DiagramStyle 数据结构
│       ├── flowchart.py           # 流程图
│       ├── funnel.py              # 漏斗
│       ├── timeline.py            # 时间线
│       ├── hierarchy.py           # 层次结构
│       ├── cycle.py               # 循环/闭环
│       ├── matrix.py              # SWOT/对比矩阵
│       ├── venn.py                # 维恩图
│       ├── pyramid.py             # 金字塔
│       └── table.py               # 表格
│
├── enterprise/                    # ★ 新增：Enterprise Pipeline
│   ├── __init__.py                # EnterprisePipeline 主流程
│   ├── scanner.py                 # ProjectScanner
│   ├── template_analyzer.py       # TemplateAnalyzer → BrandSpec
│   ├── brand_spec.py              # BrandSpec 数据结构
│   ├── enterprise_decider.py      # 受约束的 DesignDecider
│   ├── review_gate.py             # ReviewGate
│   ├── image_matcher.py           # 图片池匹配
│   ├── page_revision.py           # ★ PageRevisionEngine 页面级增删改
│   └── enterprise_renderer.py     # 基于模板的渲染器
│
├── qa/
│   └── qa_gates.py                # 不动，共享
│
└── adapters/
    └── ...                        # 不动
```

---

## 11. 实现计划

### Phase A: 基础框架（P0） ✅

1. ✅ 重构 `generate_ppt()` 为分流结构，现有逻辑提取为 `_generate_ppt_freestyle()`
2. ✅ `enterprise/` 目录结构 + `BrandSpec` 数据结构
3. ✅ `ProjectScanner` 资产扫描
4. ✅ `TemplateAnalyzer` 模板解析（含无效模板降级、srgbClr/sysClr None 值防护）
5. ✅ 入口分流逻辑（`cli.py` 新增参数，兼容 enterprise/freestyle 结果 key）
6. ✅ `EnterprisePipeline` 主流程骨架（含 dry_run/review/pages/history 全路径）

### Phase B: 核心渲染（P0） ✅

7. ✅ `EnterpriseRenderer` 基于模板渲染（含模板无效降级逻辑、`create_presentation(keep_slides=False)` 自动清除模板页）
8. ✅ LOGO 自动插入（Pillow 读取实际宽高比，避免拉伸；支持 skip_slides）
9. ✅ 模板占位符填充逻辑（`_copy_content` 跨 slide 复制标题和占位符文本）
10. ✅ `EnterpriseDesignDecider` 受约束设计决策（layout_mapping + business_mode density 建议）

### Phase C: 确认机制（P0） ✅

11. ✅ `ReviewGate` 方案展示（`generate_proposal` + `format_cli`）
12. ✅ CLI 交互确认（y/n 输入，确认后自动重新调用 generate_ppt 生成）
13. ✅ JSON 方案文件输出/读取（含错误处理，read_proposal 返回 None）

### Phase C+: 页面级修订（P0） ✅

14. ✅ `PageRevisionEngine` 页面级增删改引擎（`enterprise/page_revision.py`）— `PageOp`/`parse_pages`/`compute_target_sequence` + `revise()` 两阶段重建
15. ✅ `--pages` 参数解析（`parse_pages()` 语法解析器，含完整输入验证和错误提示）
16. ✅ 序列变换算法（`compute_target_sequence()`，含 SWAP→MOVE→DELETE→INSERT 顺序）
17. ✅ slide 删除 API（`slide_utils.remove_slide()`，共享模块，`_sldIdLst` + `drop_rel`）
18. ✅ slide 定位插入 API（两阶段重建策略：先删全部模板页，再按目标序列 add_slide）
19. ✅ slide 内容替换（`_copy_content` 跨 slide 复制标题和占位符文本）
20. ✅ meta.json `slides[]` 字段（pipeline 写入每页 goal/title）
21. ✅ 风格一致性保证（BrandSpec + layout_id + density 三层机制）
22. ✅ 修订记录（`revision_type`/`ops_applied`/`num_slides`）

### Phase D: 图片匹配（P1） ✅

23. ✅ `ImageMatcher` 文件名模糊匹配（tokenize 全词匹配，goal关键词→文件名，title兜底，轮询分配）
24. ✅ `content.json` 显式图片引用（`content_parser.py` 已实现相对路径→绝对路径转换）
25. ✅ 图片池 + ImageFetcher 降级逻辑（匹配不上时调用 FreeStyle ImageFetcher 生成/搜索图片）

### Phase E: 密集布局（P1） ✅

26. ✅ `LayoutRegistry` 新增布局 ID 12-22（grid-2x2/dense-bullets/two-column-dense/table-layout/sidebar-left/exercise-layout/code-block/timeline-horizontal/swot-matrix/funnel/cycle-diagram）
27. ✅ `density` 参数实现（DensityProfile 10级字号/间距/内容量/image比例，apply_density_to_bullets 截断，_BASELINE_IMAGE_RATIO 命名常量）
28. ✅ `StoryPlanner` 新增 education/training/report 策略 + business_mode 自动映射

### Phase F: 图形引擎（P1） ✅

29. ✅ `DiagramEngine` 主入口 + `BaseDiagram` 基类（含 connector routing + group_nodes 支持）
30. ✅ `TextMeasurer` 文字尺寸估算（CJK 宽度比 2.0 + 字体族比例）
31. ✅ `LayoutEngine` 坐标计算 + 溢出保护（`Region` frozen dataclass + subregion/inset）
32. ✅ P0 图形实现：flowchart / funnel / timeline
33. ✅ `DataSplitter` 分页拆分（nodes/stages/events 自动分页 + page_index 元数据）
34. ✅ `DiagramStyle` 样式系统（from_theme/from_brand_spec/apply_density + _color_map fallback）

### Phase G: 教学增强（P2） ✅

35. ✅ 教学动画序列（`animation.py`，12 种幻灯片切换 + 11 种入场效果，XML 注入 `<p:transition>` + `<p:timing>`）
36. ✅ P1 图形实现：swot / matrix / table / cycle
37. ✅ 代码块布局渲染（暗色背景 + Consolas 字体 + language 标签）
38. ✅ 练习/问答布局渲染（蓝色徽章 + 斜体说明 + duration 显示）

### Phase H: 高级图形（P3） ✅

39. ✅ P2 图形实现：hierarchy / pyramid / venn
40. ✅ Connector 避让算法（`connector_router.py`，线段-矩形相交检测 + 正交绕行路由）
41. ✅ 图形 add_group 打组支持（`shape_utils.py` + `BaseDiagram._group_rendered_nodes()`）

---

## 11.5 实现状态与模块清单（v3.0）

### 已实现模块

| 模块 | 文件 | 职责 | 测试数 |
|------|------|------|--------|
| `ProjectScanner` | `enterprise/scanner.py` | 项目文件夹资产扫描（启发式模板选择） | 8 |
| `ProjectAsset` | `enterprise/scanner.py` | 资产数据结构 | — |
| `BrandSpec` | `enterprise/brand_spec.py` | 品牌规范数据结构（from_brand_json/merge） | 8 |
| `TemplateAnalyzer` | `enterprise/template_analyzer.py` | 模板 XML 解析（颜色/字体/布局） | 5 |
| `PageOp`/`parse_pages`/`compute_target_sequence` | `enterprise/page_revision.py` | 页面操作解析与序列变换（多 insert 同位支持） | 30 |
| `PageRevisionEngine` | `enterprise/page_revision.py` | 两阶段重建 revise()（含 insert 定位 + dropped content 兜底） | 11 |
| `VersionManager` | `enterprise/version_manager.py` | 版本号/元数据管理 | 5 |
| `EnterpriseRenderer` | `enterprise/enterprise_renderer.py` | 模板驱动渲染 + LOGO 插入 + 智能布局选择 | 6 |
| `EnterpriseDesignDecider` | `enterprise/enterprise_decider.py` | 受约束设计决策 | 3 |
| `ReviewGate` | `enterprise/review_gate.py` | 方案展示/JSON 输出/CLI 交互 | 4 |
| `ContentParser` | `enterprise/content_parser.py` | content.json slides[] 解析（含 diagram/code/exercise） | 5 |
| `ImageMatcher` | `enterprise/image_matcher.py` | 图片池自动匹配（tokenize 全词匹配） | 9 |
| `DensityProfile` | `enterprise/density_profile.py` | 10级密度配置（字号/间距/内容量） | 9 |
| `EnterprisePipeline` | `enterprise/pipeline.py` | 主流程编排（全路径 + 密度 + cards + diagram + code + exercise + animation） | 7 |
| `slide_utils` | `enterprise/slide_utils.py` | 共享 slide 删除工具 | 3 |
| CLI | `cli.py` | 企业参数 + 结果兼容 + review 交互 | 17 |
| 入口分流 | `__init__.py` | `generate_ppt()` → freestyle/enterprise | — |
| `DiagramEngine` | `renderer/diagram_engine.py` | 图形引擎主入口（10 种图形注册 + 文本降级） | 43 |
| `BaseDiagram` | `renderer/diagram/base.py` | 图形基类（node/connector 绘制 + routing + grouping） | — |
| `DiagramStyle` | `renderer/diagram/diagram_style.py` | 图形样式系统（from_theme/from_brand_spec/apply_density） | 8 |
| `Region` | `renderer/diagram/layout_engine.py` | 坐标计算 + subregion/inset | 7 |
| `TextMeasurer` | `renderer/diagram/text_measurer.py` | 文字尺寸估算（CJK + 字体族比例） | 6 |
| `FlowchartDiagram` | `renderer/diagram/flowchart.py` | 横向/纵向流程图 | — |
| `FunnelDiagram` | `renderer/diagram/funnel.py` | 递减宽度漏斗 | — |
| `TimelineDiagram` | `renderer/diagram/timeline.py` | 交替标签时间线 | — |
| `SwotDiagram` | `renderer/diagram/swot.py` | 4 象限 SWOT 分析 | — |
| `MatrixDiagram` | `renderer/diagram/matrix.py` | 行列标签 + 单元格矩阵 | — |
| `TableDiagram` | `renderer/diagram/table.py` | 表头 + 交替行色表格 | — |
| `CycleDiagram` | `renderer/diagram/cycle.py` | 圆形排列循环图 | — |
| `HierarchyDiagram` | `renderer/diagram/hierarchy.py` | 多层级 + parent→child 连接 | — |
| `PyramidDiagram` | `renderer/diagram/pyramid.py` | 递减宽度金字塔 | — |
| `VennDiagram` | `renderer/diagram/venn.py` | 2-3 集合维恩图 | — |
| `DataSplitter` | `renderer/diagram/data_splitter.py` | 分页拆分 | 5 |
| `ConnectorRouter` | `renderer/diagram/connector_router.py` | 连线避让（线段-矩形相交 + 正交绕行） | 9 |
| `Animation` | `renderer/animation.py` | 幻灯片切换动画 + 入场效果（XML 注入） | 16 |
| `ShapeUtils` | `renderer/shape_utils.py` | 形状分组（add_group_shape） | — |

### 测试文件

| 文件 | 测试数 | 覆盖范围 |
|------|--------|---------|
| `test_enterprise.py` | 29 | 基础模块单元测试 |
| `test_enterprise_deep.py` | 31 | 边界情况 + TemplateAnalyzer + 多 insert 同位 |
| `test_enterprise_phase2.py` | 17 | CLI + pipeline + version |
| `test_enterprise_renderer.py` | 12 | EnterpriseRenderer + EnterpriseDesignDecider |
| `test_enterprise_gate_content.py` | 9 | ReviewGate + ContentParser |
| `test_page_revision_revise.py` | 10 | PageRevisionEngine.revise() 两阶段重建 |
| `test_pipeline_integration.py` | 7 | 全流程集成测试 |
| `test_audit_edge_cases.py` | 27 | 审计1修复边界测试 |
| `test_populate_slide.py` | 8 | 内容渲染测试 |
| `test_image_matcher.py` | 9 | 图片匹配测试 |
| `test_e2e_smoke.py` | 9 | 端到端冒烟测试 |
| `test_density_profile.py` | 9 | 密度配置测试 |
| `test_story_strategies.py` | 13 | StoryPlanner 新策略 + business_mode 映射 |
| `test_cards_and_layouts.py` | 16 | cards 渲染 + 高密度布局 |
| `test_audit_round2.py` | 18 | 审计2修复边界测试 |
| `test_diagram_engine.py` | 62 | 全部图形类型 + DiagramEngine + code/exercise |
| `test_animation.py` | 16 | 幻灯片切换 + 入场动画 + pipeline 集成 |
| `test_connector_router.py` | 9 | 连线避让 + 线段相交 |
| **合计** | **317** | + 58 FreeStyle/layout = **375 总测试** |

### v2.0 审计修复记录（第1轮）

| 严重度 | 问题 | 修复 |
|--------|------|------|
| CRITICAL | `PageOp.target: int = None` 类型错误 | → `int \| None = None` |
| HIGH | `parse_pages` 无输入验证 | → split 长度检查 + ValueError 捕获 |
| HIGH | `card["image"]` KeyError | → `card.get("image")` + 安全路径解析 |
| HIGH | `--pages` 无 template 时静默忽略 | → 返回明确 error 字段 |
| HIGH | Template 每页重复打开 N 次 | → 循环外一次打开，缓存 layout_names |
| HIGH | `_remove_slide` 重复实现 | → 提取到 `slide_utils.py` 共享模块 |
| HIGH | Logo `add_picture` 拉伸变形 | → Pillow 读取实际宽高比 |
| MEDIUM | `srgb.get('val')` 可能返回 `"#None"` | → None 检查 |
| MEDIUM | `TemplateAnalyzer` 无错误处理 | → try/except 返回 fallback BrandSpec |
| MEDIUM | `density=0` 被当作 falsy | → `density if density is not None` |
| MEDIUM | `animation_style` 语义错误 | → 支持 `dark_mode`/`color_scheme` 键 |
| MEDIUM | `review_gate.read_proposal` 无错误处理 | → 返回 None |
| MEDIUM | CLI `result['page_count']` KeyError | → 兼容 `num_slides`/`page_count` |
| MEDIUM | `diagram` 非 dict 时崩溃 | → `isinstance(diagram, dict)` 检查 |
| LOW | 未使用的 imports | → 清理 |
| LOW | `slide_layouts[-1]` 空 list 崩溃 | → 添加空检查 |

### v2.2 审计修复记录（第2轮）

| 严重度 | 问题 | 修复 |
|--------|------|------|
| CRITICAL | `add_slide()` 逻辑反转：`if not slide_layouts` → 访问 `[0]` | → 重写：优先选 title 布局，避免 blank |
| HIGH | `_render_cards` magic number `5` | → `MSO_SHAPE.ROUNDED_RECTANGLE` |
| HIGH | `img_h / img_w` 除零风险 | → `if img_w > 0 else 0.75` 显式 guard |
| HIGH | CLI `--pages requires --from-version` 阻塞合法用法 | → `--pages requires --project` |
| HIGH | insert 新 slide 用 blank 布局（无 title placeholder） | → 优先选含 "title" 的布局 |
| HIGH | `insert_pos` 被忽略，新 slide 总是追加到末尾 | → 合并 existing+insert 列表按序创建 |
| MEDIUM | image_matcher 子串误匹配（`task`↔`ask`, `spain`↔`pain`） | → tokenize 按 `_`/`-`/空格 拆词，全词匹配 |
| MEDIUM | `PP_PLACEHOLDER` 字符串比较 `"BODY" in ph_type` | → 枚举值直接比较 `PP_PLACEHOLDER.BODY` |
| MEDIUM | `_build_image_fetcher` 无错误处理 | → try/except 返回 None |
| MEDIUM | `brand_spec.dark_mode` 非布尔值处理（`1`/`"yes"` → False） | → `bool()` 转换 |
| MEDIUM | scanner output 目录排除不精确（只检查 parent.name） | → `entry.parent != output_dir` Path 比较 |
| MEDIUM | `0.38` 魔数无解释 | → `_BASELINE_IMAGE_RATIO = 0.38` 命名常量 |

### 已知限制

1. ~~**`_populate_slide` 仅填充标题**~~ → ✅ 已修复（subtitle/bullets/image/cards/diagram/code/exercise 已实现）
2. ~~**CLI `--review` 无交互确认**~~ → ✅ 已修复（y/n 交互已实现）
3. ~~**Image pool 未自动分配**~~ → ✅ 已修复（ImageMatcher 已实现）
4. ~~**`--from-version` 未加载历史上下文**~~ → ✅ 已修复（读取 meta.json slides[] 重建内容）
5. ~~**Scanner 模板识别过宽**~~ → ✅ 已修复（优先选择文件名含 "template" 的 .pptx）
6. ~~**`_populate_slide` cards 未渲染**~~ → ✅ 已修复（_render_cards 圆角矩形卡片 + 图片）
7. ~~**ImageFetcher 降级未集成**~~ → ✅ 已修复（pipeline 传入 image_mode/image_config）

### 残留限制（低优先级）— 全部已修复

1. ~~**`_copy_content` 布局不匹配时静默丢内容**~~ → ✅ 已修复（`_recover_dropped_content` textbox 兜底）
2. ~~**`compute_target_sequence` 多个 insert 同位置覆盖**~~ → ✅ 已修复（`dict[int, list[PageOp]]` + offset 插入）
3. ~~**`density` 语义混淆**~~ → ✅ 已修复（density 仅控制视觉密度，页数由 `max(5, min(20, density+2))` 决定）
4. ~~**`_generate_ppt_enterprise` LLM 参数未打包**~~ → ✅ 已修复（`_build_image_fetcher` 合并顶层 kwargs 到 image_config）
5. ~~**LayoutRegistry 坐标重复**~~ → ✅ 已修复（提取 `_TITLE_STD`/`_TITLE_COMPACT`/`_DIVIDER_STD` 常量）

---

## 12. DiagramEngine — 图形引擎设计

### 12.1 问题分析

企业PPT中大量使用结构化图形表达业务逻辑，当前系统**只能画矩形和圆角矩形**，完全没有图形化建模能力。

**当前能力盘点**：

| 能力 | 现状 |
|------|------|
| Shape 类型 | 仅 RECTANGLE + ROUNDED_RECTANGLE（2种） |
| 连接线 | 无（add_connector 未使用） |
| 表格 | 无（add_table 未使用） |
| 组合图形 | 无（add_group 未使用） |
| 自由路径 | 无（build_freeform 未使用） |
| SmartArt | 无（python-pptx 无原生 API） |
| 图表 | 8种基础图表（ChartBuilder） |
| 视觉效果 | 阴影/渐变背景/发光（3种） |

**企业常见图形需求**：

| 图形类型 | 示例场景 | 用户是否需要编辑 |
|---------|---------|----------------|
| 流程图/流水线 | 数据处理流水线、业务审批流程 | 是 |
| 分析模型 | SWOT、波特五力、商业画布 | 是 |
| 层次结构 | 组织架构、技术架构 | 是 |
| 循环/闭环 | PDCA、飞轮效应 | 是 |
| 漏斗 | 销售漏斗、转化漏斗 | 是 |
| 时间线 | 项目里程碑、发展历程 | 是 |
| 对比矩阵 | 竞品对比、方案对比 | 是 |
| 关系图 | 维恩图、生态图 | 是 |

### 12.2 解决方案：原生 Shape 组合绘制

**核心决策：结构化图形用原生 Shape 绘制，视觉/氛围类用图片。**

原则：
- **结构化图形**（流程图、SWOT、漏斗、时间线、组织架构等）→ 必须用 python-pptx 原生 Shape + Connector + Textbox 组合绘制，确保可编辑
- **视觉/氛围类**（背景图、概念图、氛围图、产品展示图等）→ 使用 AI 生成图片或用户提供的图片
- 判断标准：**用户是否需要编辑其中的文字和结构？** 需要编辑 → Shape；不需要编辑 → 图片

| 内容类型 | 渲染方式 | 可编辑 | 示例 |
|---------|---------|--------|------|
| 流程图/流水线 | 原生 Shape | ✓ | 数据处理流水线、审批流程 |
| 分析模型 | 原生 Shape | ✓ | SWOT、波特五力、商业画布 |
| 层次结构 | 原生 Shape | ✓ | 组织架构、技术架构 |
| 漏斗/金字塔 | 原生 Shape | ✓ | 销售漏斗、马斯洛需求 |
| 时间线 | 原生 Shape | ✓ | 项目里程碑、发展历程 |
| 循环/闭环 | 原生 Shape | ✓ | PDCA、飞轮效应 |
| 对比矩阵/表格 | 原生 Shape | ✓ | 竞品对比、方案对比 |
| 封面背景图 | AI图片/用户图片 | ✗ | 城市天际线、科技感背景 |
| 概念氛围图 | AI图片/用户图片 | ✗ | "创新"、"协作"等抽象概念 |
| 产品展示图 | 用户图片优先 | ✗ | 产品截图、实物照片 |
| 团队照片 | 用户图片 | ✗ | 团队合影、办公室 |
| 场景插图 | AI图片/用户图片 | ✗ | 使用场景、客户案例 |

**溢出降级策略**：当结构化图形内容无法在一页内完整呈现时，采用分页/简化，而非图片替代：
1. 缩小字号重试
2. 缩小间距重试
3. 切换布局方向（水平→垂直）
4. 拆分为多页（如10步流程图拆为2页各5步）
5. 简化内容（如SWOT每格只保留前2条要点，完整内容放在下一页详情页）

### 12.3 架构设计

DiagramEngine 位于 `renderer/diagram/` 目录下，与第9章目录结构一致：

```
src/ppt_pro_max/renderer/
├── diagram_engine.py          # ★ 图形引擎主入口
└── diagram/                   # ★ 图形类型实现
    ├── __init__.py
    ├── base.py                # BaseDiagram 基类
    ├── layout_engine.py       # 坐标计算、自适应排版
    ├── text_measurer.py       # 文字尺寸估算
    ├── data_splitter.py       # 分页拆分
    ├── diagram_style.py       # DiagramStyle 数据结构
    ├── flowchart.py           # 流程图
    ├── funnel.py              # 漏斗
    ├── timeline.py            # 时间线
    ├── hierarchy.py           # 层次结构
    ├── cycle.py               # 循环/闭环
    ├── matrix.py              # SWOT/对比矩阵
    ├── venn.py                # 维恩图
    ├── pyramid.py             # 金字塔
    └── table.py               # 表格
```

### 12.4 核心接口

```python
class DiagramEngine:
    """图形引擎：用原生 Shape 组合绘制结构化图形"""

    # 注册的图形类型
    _registry: dict[str, type[BaseDiagram]] = {}

    def render(self, slide, diagram_type: str, data: dict,
               style: DiagramStyle, region: Region) -> None:
        """
        slide: python-pptx slide 对象
        diagram_type: "flowchart" | "funnel" | "timeline" | ...
        data: 图形数据（节点、连线、标签等）
        style: 颜色/边框/阴影/字号等样式
        region: (left, top, width, height) 绘制区域（英寸）
        """
        diagram_cls = self._registry.get(diagram_type)
        if diagram_cls is None:
            # 不支持的图形类型 → 文字列表替代（保持可编辑）
            self._render_as_text_list(slide, diagram_type, data, style, region)
            return
        diagram = diagram_cls(data, style, region)
        diagram.render(slide)
```

### 12.5 排版引擎：核心难点

图形绘制最大的挑战不是"画什么形状"，而是**如何在给定区域内自适应排版**——节点数量不固定、文字长度不固定、间距要合理。

#### 12.5.1 Region 概念

所有图形都在一个 Region 内绘制，Region 由 Layout 分配：

```python
@dataclass
class Region:
    left: float      # 英寸
    top: float
    width: float
    height: float

    @property
    def center_x(self) -> float:
        return self.left + self.width / 2

    @property
    def center_y(self) -> float:
        return self.top + self.height / 2

    def subregion(self, left_pct, top_pct, width_pct, height_pct) -> "Region":
        """按百分比切分子区域"""
        return Region(
            left=self.left + self.width * left_pct,
            top=self.top + self.height * top_pct,
            width=self.width * width_pct,
            height=self.height * height_pct,
        )
```

#### 12.5.2 自适应排版算法

**核心问题**：N 个节点 + M 条连线，在 W×H 的区域内如何排列？

**方案：两阶段排版**

```
阶段1: 布局计算（纯数学，不涉及 python-pptx）
  → 输出每个节点的 (x, y, width, height) 和每条连线的路径

阶段2: 渲染执行（调用 python-pptx API）
  → 按计算结果创建 Shape、Connector、Textbox
```

**阶段1 的关键：文字测量**

节点大小取决于文字长度。python-pptx 没有文字测量 API，需要预估：

```python
class TextMeasurer:
    """估算文字在 PPT 中占用的尺寸"""

    # 经验值：不同字号下每个字符的平均宽度（英寸）
    # 基于 Calibri/Arial 的测量，其他字体按比例调整
    # 系数 0.006（偏保守估算，确保不溢出；远期用 Pillow render 精确测量）
    CHAR_WIDTH_TABLE = {
        10: 0.06,   # 10pt 字号，每字符约 0.06 英寸宽
        11: 0.066,
        12: 0.072,
        13: 0.078,
        14: 0.084,
        16: 0.096,
        18: 0.108,
        20: 0.12,
        24: 0.144,
        28: 0.168,
        32: 0.192,
    }

    # CJK 字符宽度约为西文的 2 倍
    CJK_WIDTH_RATIO = 2.0

    @classmethod
    def estimate_text_size(cls, text: str, font_size_pt: int,
                           max_width: float, font_family: str = "Calibri") -> tuple[float, float]:
        """
        估算文字在 PPT 中的尺寸
        返回: (actual_width, actual_height) 英寸

        算法:
        1. 区分 CJK 和西文字符，分别计算宽度
        2. 考虑 word_wrap：超过 max_width 时换行
        3. 行高 = font_size_pt / 72 * line_spacing
        4. 加上 textbox 的内边距 (margin_left/right/top/bottom)
        """
        char_width = cls.CHAR_WIDTH_TABLE.get(font_size_pt, font_size_pt * 0.006)

        # 字体修正系数
        font_ratio = 1.0
        if "serif" in font_family.lower() or "Georgia" in font_family:
            font_ratio = 1.1
        elif "mono" in font_family.lower() or "Consolas" in font_family:
            font_ratio = 1.2

        # 逐字符计算宽度
        total_width = 0
        for ch in text:
            if cls._is_cjk(ch):
                total_width += char_width * cls.CJK_WIDTH_RATIO * font_ratio
            else:
                total_width += char_width * font_ratio

        # 换行计算
        line_height = font_size_pt / 72 * 1.2  # 1.2x 行距
        if total_width <= max_width:
            lines = 1
        else:
            lines = math.ceil(total_width / max_width)

        # textbox 内边距
        padding_v = 0.08  # 上下各 0.08 英寸
        padding_h = 0.15  # 左右各 0.15 英寸

        actual_width = min(total_width + padding_h * 2, max_width)
        actual_height = lines * line_height + padding_v * 2

        return actual_width, actual_height
```

#### 12.5.3 各图形类型的排版策略

**流程图（Flowchart）**

```
排版策略：水平流水线 或 垂直流程

水平流水线（节点 ≤ 6）:
  ┌──────┐   ┌──────┐   ┌──────┐   ┌──────┐
  │ Step1│──→│ Step2│──→│ Step3│──→│ Step4│
  └──────┘   └──────┘   └──────┘   └──────┘

  节点宽度 = (region.width - (n-1) * gap) / n
  gap = min(0.4, region.width * 0.03)
  节点高度 = region.height * 0.35
  节点 y = region.center_y - 节点高度/2

垂直流程（节点 > 6 或文字较长）:
  ┌──────┐
  │ Step1│
  └──┬───┘
     ↓
  ┌──────┐
  │ Step2│
  └──┬───┘
     ↓
  ┌──────┐
  │ Step3│
  └──────┘

  节点高度 = (region.height - (n-1) * gap) / n
  gap = min(0.3, region.height * 0.03)
  节点宽度 = region.width * 0.6
  节点 x = region.center_x - 节点宽度/2

分支流程（有条件分支）:
  ┌──────┐
  │ Step1│
  └──┬───┘
     ↓
  ┌──┴───┐
  │Decision│
  └─┬──┬─┘
   ↓    ↓
  Yes   No
   ↓    ↓
  S2    S3

  使用层级布局：每层节点水平居中分布
  层间距 = (region.height - layers * node_h) / (layers - 1)
```

**漏斗（Funnel）**

```
排版策略：等间距递减宽度

  ┌────────────────────────────┐  stage 1: width = 100%
  │       潜在客户 10,000      │
  ├──────────────────────┤      stage 2: width = 75%
  │     意向客户 3,000    │
  ├────────────────┤            stage 3: width = 50%
  │   谈判中 800   │
  ├──────────┤                stage 4: width = 30%
  │ 成交 200 │
  └──────────┘

  每层高度 = (region.height - (n-1) * gap) / n
  每层宽度 = region.width * (1 - stage_index / (n - 1) * 0.7)
  每层 x = region.center_x - width / 2
  使用 MSO_SHAPE.TRAPEZOID 或 手动绘制梯形
```

**时间线（Timeline）**

```
排版策略：水平轴线 + 交替上下标注

  ─────●──────────●──────────●──────────●──────
       │          │          │          │
    2024 Q1    2024 Q3    2025 Q1    2025 Q4
    产品立项    MVP发布    种子轮     正上线
       │          │          │          │

  轴线 y = region.center_y
  节点间距 = region.width / (n + 1)
  标注交替在轴线上方/下方
  上方标注: text_y = axis_y - 0.3 - text_height
  下方标注: text_y = axis_y + 0.3
  日期字号: 11pt, 标签字号: 13pt
```

**SWOT 矩阵（Matrix）**

```
排版策略：2×2 等分网格

  ┌──────────────┬──────────────┐
  │   S 优势     │   W 劣势     │
  │  ─────────   │  ─────────   │
  │  • 技术领先  │  • 品牌低    │
  │  • 团队强    │  • 资金少    │
  ├──────────────┼──────────────┤
  │   O 机会     │   T 威胁     │
  │  ─────────   │  ─────────   │
  │  • 市场增长  │  • 巨头入场  │
  │  • 政策支持  │  • 监管趋严  │
  └──────────────┴──────────────┘

  cell_width = region.width / 2
  cell_height = region.height / 2
  标题字号: 16pt bold, 内容字号: 12pt
  每个单元格内边距: 0.2"
  S/W 用主色背景, O/T 用辅色背景
```

**循环/闭环（Cycle）**

```
排版策略：圆形分布 + 弧形箭头

        ┌───┐
       │Plan │
    ┌──┘     └──┐
    │           │
  Act         Do
    │           │
    └──┐     ┌──┘
       │Check│
        └───┘

  节点均匀分布在圆周上
  圆心 = region.center
  半径 = min(region.width, region.height) * 0.35
  节点角度 = 360° / n * i
  节点 x = center_x + radius * cos(angle)
  节点 y = center_y + radius * sin(angle)
  连线：弧形箭头（用 build_freeform + add_line_segments 多段折线近似弧形，无原生 Bezier 支持）
  中心标签字号: 20pt, 节点标签字号: 14pt
```

**层次结构（Hierarchy）**

```
排版策略：树形层级布局

          ┌─────────┐
          │  CEO    │
          └────┬────┘
          ┌────┴────────┐
     ┌────┴───┐    ┌────┴───┐
     │  CTO   │    │  CFO   │
     └──┬─────┘    └────────┘
    ┌───┴───┐
    │       │
  Dev1    Dev2

  每层节点水平均匀分布
  层间距 = region.height / (max_depth + 1)
  同层节点间距 = region.width / (同层节点数 + 1)
  连线：父节点底部中心 → 子节点顶部中心
```

**金字塔（Pyramid）**

```
排版策略：等间距递增宽度

      ┌───┐          level 1: width = 30%
     │顶层 │
    ┌┴─────┴┐        level 2: width = 55%
   │ 中层   │
  ┌┴─────────┴┐      level 3: width = 80%
 │   基层     │
 └─────────────┘

  每层高度 = (region.height - (n-1) * gap) / n
  每层宽度 = region.width * (0.3 + 0.5 * level_index / (n - 1))
  使用 MSO_SHAPE.TRAPEZOID
```

**表格（Table）**

```
排版策略：python-pptx add_table

  ┌────────┬────────┬────────┐
  │ 方案A  │ 方案B  │ 方案C  │
  ├────────┼────────┼────────┤
  │ 价格   │ ¥99    │ ¥199   │
  │ 功能   │ 基础   │ 全部   │
  └────────┴────────┴────────┘

  使用 slide.shapes.add_table(rows, cols, left, top, width, height)
  表头行：主色背景 + 白色文字 14pt bold
  数据行：交替 muted/white 背景 + 前景色文字 12pt
  边框：1pt border color
```

### 12.6 DiagramStyle — 图形样式系统

图形样式需要与 PPT 主题一致，同时支持企业品牌规范覆盖：

```python
@dataclass
class DiagramStyle:
    # 节点样式
    node_fill: str = "primary"          # 填充色角色 (primary/secondary/accent/muted)
    node_border: str = "border"         # 边框色角色
    node_border_width_pt: float = 1.0
    node_corner_radius: str = "rounded" # "sharp" | "rounded" | "pill"
    node_shadow: bool = False

    # 节点文字样式
    node_font_size_pt: int = 13
    node_font_weight: str = "normal"    # "normal" | "bold"
    node_font_color: str = "on-primary" # 文字色角色
    node_text_alignment: str = "center" # "left" | "center"

    # 连线样式
    connector_color: str = "muted-foreground"
    connector_width_pt: float = 1.5
    connector_style: str = "solid"      # "solid" | "dashed"
    arrow_enabled: bool = True
    arrow_size: str = "medium"          # "small" | "medium" | "large"

    # 标签样式（时间线日期、漏斗数值等辅助文字）
    label_font_size_pt: int = 11
    label_font_color: str = "muted-foreground"

    # 间距
    node_gap_inches: float = 0.3        # 节点间距
    padding_inches: float = 0.15        # 节点内边距

    # 特殊：SWOT 等矩阵的单元格样式
    cell_header_font_size_pt: int = 16
    cell_body_font_size_pt: int = 12
    cell_header_font_weight: str = "bold"

    @classmethod
    def from_theme(cls, theme: dict, business_mode: str = "pitch") -> "DiagramStyle":
        """从 PPT 主题生成默认图形样式"""
        style = cls()

        # 暗色模式适配
        bg = theme.get("colors", {}).get("background", "#FFFFFF")
        is_dark = cls._is_dark(bg)
        if is_dark:
            style.node_fill = "muted"
            style.node_font_color = "foreground"
            style.connector_color = "muted-foreground"

        # business_mode 适配
        if business_mode in ("education", "training"):
            style.node_font_size_pt = 12      # 教育场景字号偏小
            style.label_font_size_pt = 10
            style.node_gap_inches = 0.2       # 间距紧凑
            style.node_shadow = False          # 去掉装饰性效果

        return style

    @classmethod
    def from_brand_spec(cls, brand_spec: "BrandSpec") -> "DiagramStyle":
        """从企业品牌规范生成图形样式"""
        style = cls()
        if brand_spec.spacing:
            style.node_font_size_pt = brand_spec.spacing.get("body_size_pt", 13)
            style.node_gap_inches = brand_spec.spacing.get("margins_inches", 0.3) * 0.3
        return style
```

### 12.7 字号体系 — 密度与层级的平衡

图形中的文字不是孤立的，需要与页面标题、正文形成清晰的层级关系：

```
字号层级体系（以 16:9 宽屏为基准）：

页面标题:     28pt bold    ← 现有 LayoutRegistry 定义
页面副标题:   18pt normal  ← 现有
───────────────────────────────────
图形节点标题: 13-14pt bold ← DiagramStyle.node_font_size_pt
图形节点内容: 11-12pt      ← DiagramStyle.label_font_size_pt
图形辅助标签: 10-11pt      ← 日期、数值等
───────────────────────────────────
页面正文:     14-16pt      ← 现有 body font_size
页面 bullet:  14-16pt      ← 现有
```

**density 对图形字号的影响**：

| density | 节点标题 | 节点内容 | 辅助标签 | 节点间距 | 节点内边距 |
|---------|---------|---------|---------|---------|-----------|
| 1-3     | 16pt    | 13pt    | 12pt    | 0.4"    | 0.2"      |
| 4-6     | 14pt    | 12pt    | 11pt    | 0.3"    | 0.15"     |
| 7-10    | 12pt    | 10pt    | 9pt     | 0.2"    | 0.1"      |

**关键约束**：图形内字号不得大于页面正文，否则视觉层级混乱。

### 12.8 坐标计算 — 自适应与溢出保护

#### 12.8.1 计算流程

```
输入: diagram_data + diagram_style + region

1. 预测量阶段
   for each node in data.nodes:
       node.text_size = TextMeasurer.estimate_text_size(
           node.label, style.node_font_size_pt,
           max_width=region.width * 0.8  # 单节点最大不超过区域80%
       )

2. 布局计算阶段
   layout = LayoutEngine.calculate(diagram_type, nodes, edges, region, style)
   → 输出: {node_id: (x, y, w, h), edge_id: path}

3. 溢出检测
   if layout.overflows(region):
       # 策略1: 缩小字号重试
       style.node_font_size_pt -= 2
       goto 1
       # 策略2: 缩小间距重试
       style.node_gap_inches *= 0.8
       goto 2
       # 策略3: 切换布局方向（水平→垂直）
       layout = LayoutEngine.calculate(alternate_type, ...)
       # 策略4: 拆分为多页
       split_data = DataSplitter.split(data, max_nodes_per_page)
       → 生成多张 slide，每张可编辑
       # 策略5: 简化内容（保留核心，完整内容放下一页）
       simplified_data = DataSimplifier.trim(data, max_items=3)

4. 渲染执行
   for each node: slide.shapes.add_shape(...)
   for each edge: slide.shapes.add_connector(...)
```

#### 12.8.2 溢出保护

```python
class LayoutEngine:
    MAX_RETRIES = 3

    def calculate(self, diagram_type, nodes, edges, region, style) -> LayoutResult:
        for attempt in range(self.MAX_RETRIES):
            result = self._do_calculate(diagram_type, nodes, edges, region, style)
            if not result.overflows(region):
                return result

            # 自适应调整
            if attempt == 0:
                # 第一次溢出：缩小字号
                style = style.copy(node_font_size_pt=max(style.node_font_size_pt - 2, 9))
            elif attempt == 1:
                # 第二次溢出：缩小间距
                style = style.copy(node_gap_inches=style.node_gap_inches * 0.7)
            else:
                # 第三次溢出：拆分为多页
                result.needs_split = True

        return result
```

#### 12.8.3 最小字号保护

```
节点标题最小: 9pt（低于此值无法阅读）
辅助标签最小: 8pt
如果缩到最小仍溢出 → 拆分为多页（每页保持可编辑）
```

### 12.9 Connector（连线）实现

python-pptx 的 `add_connector` 支持直线、折线、曲线三种：

```python
from pptx.enum.shapes import MSO_CONNECTOR_TYPE

# 直线连接
connector = slide.shapes.add_connector(
    MSO_CONNECTOR_TYPE.STRAIGHT,   # 直线
    start_x, start_y,              # 起点（英寸）
    end_x, end_y,                  # 终点（英寸）
)

# 折线连接（用于流程图拐弯）
connector = slide.shapes.add_connector(
    MSO_CONNECTOR_TYPE.ELBOW,      # 折线
    start_x, start_y,
    end_x, end_y,
)

# 样式设置
connector.line.color.rgb = RGBColor.from_string(color.lstrip("#"))
connector.line.width = Pt(width_pt)

# 虚线
connector.line.dash_style = MSO_LINE.DASH

# 箭头（通过 XML）
# 注意：a:ln 是 p:spPr 的子元素，不是 p:cxnSp 的直接子元素
# connector._element.find(qn("a:ln")) 不会递归搜索，需通过 spPr 中转
spPr = connector._element.find(qn("p:spPr"))
line_xml = spPr.find(qn("a:ln")) if spPr is not None else connector._element.find(".//" + qn("a:ln"))
tail_arrow = etree.SubElement(line_xml, qn("a:tailEnd"))
tail_arrow.set("type", "triangle")
tail_arrow.set("w", "med")
tail_arrow.set("len", "med")
```

### 12.10 Shape 类型使用计划

| 图形类型 | 使用的 MSO_SHAPE | 说明 |
|---------|-----------------|------|
| 流程图节点 | RECTANGLE / ROUNDED_RECTANGLE | 普通步骤用圆角矩形 |
| 流程图判断 | DIAMOND | 菱形表示条件分支 |
| 流程图起止 | OVAL | 椭圆表示开始/结束 |
| 漏斗层 | TRAPEZOID | 梯形 |
| 时间线节点 | OVAL | 圆点 |
| 时间线轴线 | RECTANGLE | 细长矩形 |
| 循环节点 | OVAL / ROUNDED_RECTANGLE | 圆形或圆角矩形 |
| SWOT 单元格 | RECTANGLE | 矩形 |
| 金字塔层 | TRAPEZOID | 梯形 |
| 箭头 | RIGHT_ARROW / CHEVRON | 方向指示 |

> **圆角控制**：`ROUNDED_RECTANGLE` 的圆角半径通过 `shape.adjustments[0]` 控制（默认约 0.167）。DiagramStyle 的 `node_corner_radius` 映射：`"sharp"` → `adjustments[0] = 0.02`，`"rounded"` → `adjustments[0] = 0.15`（默认），`"pill"` → `adjustments[0] = 0.5`。

> **文字垂直居中**：`shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER` 仅实现水平居中。垂直居中需通过 XML：`text_frame._element.find(qn('a:bodyPr')).set('anchor', 'ctr')`。

> **连线宽度**：`connector.line.width` 只接受整数 EMU，不接受 float。必须使用 `connector.line.width = Pt(width_pt)` 转换，不可直接赋 float 值。

> **打组约束**：`add_group_shape(shapes_iterable)` 要求 shapes 已添加到 slide 上。实现顺序：1) add 所有节点/连线到 slide → 2) 收集 shape 对象 → 3) 调用 `add_group_shape()`。

> **Connection site 索引**：`ROUNDED_RECTANGLE` 有 4 个 connection site（索引 0-3），对应方向：0=顶部，1=左侧，2=底部，3=右侧。`begin_connect(shape, site_index)` 中 site_index 超出范围会抛 `KeyError`。流程图连线时需根据节点相对位置选择正确的 site index。

### 12.11 与 Pipeline 集成

#### LayoutRegistry 扩展

```python
# 布局 ID 20: Diagram Focus
"diagram-focus": {
    "id": 20,
    "name": "Diagram Focus",
    "goal_mapping": ["process", "pipeline", "funnel", "timeline",
                     "hierarchy", "cycle", "swot", "comparison",
                     "pyramid", "architecture"],
    "placeholders": {
        "title": {
            "x": 0.9, "y": 0.5, "width": 11.533, "height": 0.7,
            "font_size": 28, "font_weight": "bold", "alignment": "left",
            "color_role": "foreground", "font_role": "heading",
        },
        "diagram": {
            "x": 0.9, "y": 1.5, "width": 11.533, "height": 5.4,
            "type": "diagram",
        },
        "insight": {
            "x": 0.9, "y": 7.0, "width": 11.533, "height": 0.4,
            "font_size": 11, "font_weight": "normal", "alignment": "left",
            "color_role": "muted-foreground", "font_role": "body",
        },
    },
}

# 布局 ID 21: Diagram + Text（左图右文）
"diagram-plus-text": {
    "id": 21,
    "name": "Diagram + Text",
    "goal_mapping": ["process-detail", "architecture-detail"],
    "placeholders": {
        "title": { ... },
        "diagram": {
            "x": 0.9, "y": 1.5, "width": 7.5, "height": 5.4,
            "type": "diagram",
        },
        "body": {
            "x": 8.8, "y": 1.5, "width": 3.633, "height": 5.4,
            "font_size": 14, ...,
        },
    },
}

# 布局 ID 22: Table Focus
"table-focus": {
    "id": 22,
    "name": "Table Focus",
    "goal_mapping": ["data-table", "specification"],
    "placeholders": {
        "title": { ... },
        "table": {
            "x": 0.9, "y": 1.5, "width": 11.533, "height": 5.4,
            "type": "table",
        },
    },
}
```

#### ContentGenerator 扩展

```python
# PageContent 新增字段
@dataclass
class PageContent:
    # ... 现有字段 ...
    image: str | None = None           # 显式指定图片文件名（项目文件夹相对路径）
    cards: list[dict] | None = None   # 卡片数据 [{"title": ..., "body": ..., "image": ...}]
    diagram_type: str | None = None       # "flowchart" | "funnel" | ...
    diagram_data: dict | None = None      # 图形数据（含 type 字段，与 diagram_type 冗余但便于 DataSplitter 访问）
```

#### EnterprisePageDesign 扩展

Enterprise Pipeline 的设计决策需要记录模板布局索引，现有 `PageDesign` 不含此字段。方案：创建 `EnterprisePageDesign` 子类，不修改现有 `PageDesign`：

```python
@dataclass
class EnterprisePageDesign(PageDesign):
    template_layout_index: int | None = None  # 映射到 slide_layouts[index]
    template_placeholders: dict | None = None  # 该布局的占位符信息
    density_overrides: dict | None = None      # 密度调整覆盖
```

EnterpriseRenderer 根据 `template_layout_index` 选择模板布局，根据 `template_placeholders` 决定哪些内容填入占位符、哪些手动添加 shape。

> **与现有 `--content` 参数的关系**：当前 CLI 有 `--content` 参数指向一个 JSON 文件，其解析逻辑（`_load_user_content()`）将 JSON 扁平化为 context dict，无法识别 `slides[]` 结构。Enterprise Pipeline 的 `content.json` 使用新的结构化解析逻辑（按 `slides[]` 逐页映射到 `PageContent`）。两者同时存在时，`--content` 优先（用户显式指定优先于项目文件夹中的文件）。

#### content.json 解析实现方案

Enterprise Pipeline 需要新增 `_load_enterprise_content()` 方法，与现有 `_load_user_content()` 完全独立：

```python
def _load_enterprise_content(content_raw: dict, project_dir: str) -> list[PageContent]:
    """
    解析 content.json 的 slides[] 结构，逐页映射到 PageContent。
    与 _load_user_content() 的扁平化逻辑完全独立。

    content.json 结构:
    {
        "meta": {"title": ..., "subtitle": ..., "author": ..., "date": ...},
        "slides": [
            {"goal": "hook", "title": "...", "subtitle": "...", "image": "hero.png"},
            {"goal": "problem", "title": "...", "bullets": [...]},
            {"goal": "features", "title": "...", "cards": [...]},
            {"goal": "process", "title": "...", "diagram": {"type": "flowchart", ...}}
        ]
    }
    """
    slides = content_raw.get("slides", [])
    meta = content_raw.get("meta", {})
    result = []

    for i, slide_data in enumerate(slides):
        # 图片路径解析：相对路径 → 绝对路径
        image = slide_data.get("image")
        if image and not os.path.isabs(image):
            image = os.path.join(project_dir, image)

        # 卡片中的图片路径解析
        cards = slide_data.get("cards")
        if cards:
            cards = [
                {**card, "image": os.path.join(project_dir, card["image"]) if card.get("image") and not os.path.isabs(card["image"]) else card.get("image")}
                for card in cards
            ]

        # 图形数据提取
        diagram = slide_data.get("diagram")
        diagram_type = diagram.get("type") if diagram else None
        diagram_data = diagram if diagram else None

        # 第一页使用 meta 中的 title/subtitle 作为默认值
        title = slide_data.get("title", meta.get("title", "") if i == 0 else "")
        subtitle = slide_data.get("subtitle", meta.get("subtitle", "") if i == 0 else None)

        result.append(PageContent(
            goal=slide_data.get("goal", "content"),
            title=title,
            subtitle=subtitle,
            bullets=slide_data.get("bullets"),
            image=image,
            cards=cards,
            diagram_type=diagram_type,
            diagram_data=diagram_data,
        ))

    return result
```

**调用时机**：EnterprisePipeline 在 Phase 3（ContentGenerator）之前检查 `ProjectAsset.content_raw`，如果存在则调用 `_load_enterprise_content()` 直接生成 `list[PageContent]`，跳过 ContentGenerator 的自动生成逻辑。如果不存在，则走 ContentGenerator 自动生成（与 FreeStyle 相同）。

#### content.json 中的图形数据

```json
{
    "goal": "process",
    "title": "数据处理流水线",
    "diagram": {
        "type": "flowchart",
        "direction": "horizontal",
        "nodes": [
            {"id": 1, "label": "数据采集", "type": "process"},
            {"id": 2, "label": "数据清洗", "type": "process"},
            {"id": 3, "label": "特征工程", "type": "process"},
            {"id": 4, "label": "模型训练", "type": "process"},
            {"id": 5, "label": "模型部署", "type": "process"}
        ],
        "edges": [
            {"from": 1, "to": 2},
            {"from": 2, "to": 3},
            {"from": 3, "to": 4},
            {"from": 4, "to": 5}
        ]
    }
}
```

```json
{
    "goal": "swot",
    "title": "SWOT 分析",
    "diagram": {
        "type": "swot",
        "strengths": ["技术领先", "团队经验丰富", "专利壁垒"],
        "weaknesses": ["品牌知名度低", "资金有限"],
        "opportunities": ["市场年增长40%", "政策支持AI产业"],
        "threats": ["巨头入场", "监管趋严", "人才竞争"]
    }
}
```

#### EnterpriseDesignDecider 映射

```python
GOAL_TO_DIAGRAM = {
    "process": "flowchart",
    "pipeline": "flowchart",
    "funnel": "funnel",
    "timeline": "timeline",
    "milestone": "timeline",
    "hierarchy": "hierarchy",
    "org-chart": "hierarchy",
    "cycle": "cycle",
    "pdca": "cycle",
    "swot": "swot",
    "comparison": "matrix",
    "pyramid": "pyramid",
    "architecture": "hierarchy",
}
```

### 12.12 降级策略

```
1. diagram_data 完整 + 类型在支持列表 → DiagramEngine 原生绘制（可编辑）
2. diagram_data 只有 type 无 data → DiagramEngine 用默认模板数据绘制
3. 溢出重试3次仍失败 → 拆分为多页（每页可编辑）
4. diagram_type 不在支持列表 → 文字列表替代（可编辑，最朴素但永远可用）
5. 内容过多无法简化 → 分页展示（概览页 + 详情页）

**注意**：以上降级仅针对结构化图形。背景图、概念图、产品图等视觉/氛围类内容仍使用 AI 生成图片或用户图片，走 ImageFetcher 现有流程。
```

### 12.13 分页拆分策略

当图形内容无法在一页内完整呈现时，采用分页拆分，**绝不使用图片替代**。

#### 拆分规则

| 图形类型 | 拆分方式 | 示例 |
|---------|---------|------|
| flowchart | 按步骤数拆分，每页 ≤ 6 步 | 10步流程 → 第1页1-5步 + 第2页6-10步 |
| funnel | 不拆分（漏斗必须完整展示） | 溢出时缩小字号/简化标签 |
| timeline | 按时间段拆分 | 2020-2030 → 第1页2020-2025 + 第2页2025-2030 |
| hierarchy | 按层级拆分 | 第1页顶层架构 + 第2页各子部门详情 |
| swot | 不拆分（4格必须同页） | 溢出时每格精简要点数 |
| cycle | 不拆分（闭环必须完整） | 溢出时精简节点描述 |
| pyramid | 不拆分（金字塔必须完整） | 溢出时精简层级描述 |
| matrix | 按行数拆分 | 10行对比表 → 第1页1-5行 + 第2页6-10行 |

#### 拆分页的视觉衔接

```
第1页（概览/前半部分）:
  标题: "数据处理流水线 (1/2)"
  图形: Step1 → Step2 → Step3 → Step4 → Step5
  右下角提示: "→ 续下页"

第2页（详情/后半部分）:
  标题: "数据处理流水线 (2/2)"
  图形: Step6 → Step7 → Step8 → Step9 → Step10
  左上角提示: "← 接上页"
```

#### DataSplitter 实现

```python
class DataSplitter:
    """将过大的图形数据拆分为多页"""

    # 各类型每页最大节点数
    MAX_PER_PAGE = {
        "flowchart": 6,
        "timeline": 8,
        "hierarchy": 7,     # 每页最多展示7个节点
        "matrix": 6,        # 每页最多6行
        "funnel": 99,       # 不拆分
        "swot": 99,         # 不拆分
        "cycle": 99,        # 不拆分
        "pyramid": 99,      # 不拆分
    }

    @classmethod
    def split(cls, diagram_type: str, data: dict) -> list[dict]:
        """
        将图形数据拆分为多页数据
        返回: [page1_data, page2_data, ...]
        每个元素都是完整的 diagram_data，可独立渲染
        """
        max_nodes = cls.MAX_PER_PAGE.get(diagram_type, 6)

        if diagram_type == "flowchart":
            return cls._split_flowchart(data, max_nodes)
        elif diagram_type == "timeline":
            return cls._split_timeline(data, max_nodes)
        elif diagram_type == "hierarchy":
            return cls._split_hierarchy(data, max_nodes)
        elif diagram_type == "matrix":
            return cls._split_matrix(data, max_nodes)
        else:
            # 不可拆分的类型：精简内容
            return [cls._simplify(data, max_items=4)]

    @classmethod
    def _split_flowchart(cls, data: dict, max_nodes: int) -> list[dict]:
        nodes = data["nodes"]
        edges = data["edges"]
        if len(nodes) <= max_nodes:
            return [data]

        pages = []
        for i in range(0, len(nodes), max_nodes):
            chunk_nodes = nodes[i:i + max_nodes]
            chunk_ids = {n["id"] for n in chunk_nodes}
            # 只保留本页节点之间的边 + 跨页连接标记
            chunk_edges = [
                e for e in edges
                if e["from"] in chunk_ids and e["to"] in chunk_ids
            ]
            # 如果不是最后一页，添加跨页连接提示
            page_data = {
                "type": data["type"],
                "direction": data.get("direction", "horizontal"),
                "nodes": chunk_nodes,
                "edges": chunk_edges,
                "continuation": {
                    "page": i // max_nodes + 1,
                    "total_pages": (len(nodes) + max_nodes - 1) // max_nodes,
                    "has_next": i + max_nodes < len(nodes),
                    "has_prev": i > 0,
                }
            }
            pages.append(page_data)
        return pages
```

#### 与 Pipeline 集成

DiagramEngine 的分页结果需要反馈给 EnterprisePipeline，在 StoryPlan 中插入额外页面：

```python
# EnterprisePipeline 中的处理逻辑
page_contents = ContentGenerator.generate(...)

# 检查图形是否需要分页
final_contents = []
for content in page_contents:
    if content.diagram_data:
        splits = DataSplitter.split(content.diagram_type, content.diagram_data)
        if len(splits) == 1:
            final_contents.append(content)
        else:
            # 第1页保留原 goal，后续页 goal 加 "-continued"
            for i, split_data in enumerate(splits):
                continued = content.copy()
                continued.diagram_data = split_data
                continued.title = f"{content.title} ({i+1}/{len(splits)})"
                if i > 0:
                    continued.goal = f"{content.goal}-continued"
                final_contents.append(continued)
    else:
        final_contents.append(content)

# 同步更新 page_designs
page_designs = EnterpriseDesignDecider.decide(story_plan, brand_spec, ...)
# ... 根据 final_contents 的页数调整 page_designs
```

### 12.14 实现优先级

| 优先级 | 图形类型 | 理由 |
|--------|---------|------|
| P0 | flowchart | 企业最常见需求，流水线/流程 |
| P0 | funnel | 销售/营销核心图形 |
| P0 | timeline | 几乎每个 pitch 都有 |
| P1 | swot/matrix | 战略分析标配 |
| P1 | table | 数据对比必备 |
| P1 | cycle | PDCA/飞轮常见 |
| P2 | hierarchy | 组织架构/技术架构 |
| P2 | pyramid | 层次模型 |
| P2 | venn | 关系图 |

---

## 13. 风险与待决事项

| 风险 | 影响 | 缓解措施 |
|------|------|---------|
| 模板占位符格式千奇百怪 | 填充失败 | 降级为手动 shape 放置 |
| 模板布局名称无法自动匹配 | 映射错误 | brand.json 显式映射 + ReviewGate 确认 |
| 图片池模糊匹配误判 | 图片放错位置 | ReviewGate 展示匹配结果供确认 |
| 高密度布局在模板中空间不足 | 内容溢出 | 溢出检测 + 自动降级密度 |
| 教育策略的 content 生成质量 | 内容不专业 | 依赖 content.json 提供真实内容 |
| **TextMeasurer 估算不准** | **节点大小偏差** | **保守估算 + 溢出重试；远期用 Pillow 精确测量** |
| **Connector 路径穿越节点** | **图形混乱** | **Elbow connector 自动避让；或手动计算中间点** |
| **CJK 文字宽度估算偏差大** | **中文排版溢出** | **CJK 字符按 2x 宽度计算 + 额外 10% 安全边距** |
| **图形在模板布局中空间不足** | **与模板元素重叠** | **Region 从模板占位符位置推导；溢出时拆分多页** |
| **内容过多一页放不下** | **信息丢失** | **分页策略：概览页+详情页，每页保持可编辑** |
| **AI图片与品牌风格不一致** | **视觉违和** | **ImageFetcher prompt 注入品牌色/风格关键词；优先使用用户图片** |
| **模板文件损坏或路径无效** | **Presentation() 抛异常** | **EnterpriseRenderer 降级为空白 Presentation() 创建** |
| **页面级修订风格不一致** | **新页面与旧页面视觉割裂** | **BrandSpec + layout name 匹配 + density 三层保证（见 §9.6.6）** |
| **python-pptx partname 冲突** | **删除 slide 后 add_slide 复用已删除 partname，保存后数据损坏** | **两阶段重建策略：Pass1 修改内容，Pass2 从模板新建 prs 按序列复制（见 §9.6.3，已验证）** |
| **slide 删除后索引偏移** | **操作错位** | **序列变换算法：所有 page 编号指原始文档，compute_target_sequence 在原始序列上操作（见 §9.6.4）** |
| **slide_master 无 theme 属性** | **无法提取品牌色/字体** | **必须通过 XML 解析 theme part（已验证：通过 prs.part.part_related_by(RT.THEME) 获取，用 etree.fromstring(blob) 解析）** |
| **conn.line.width 不接受 float** | **TypeError** | **必须使用 Pt() 转换** |
| **connector a:ln 非顶层元素** | **find() 返回 None** | **通过 p:spPr 中转或 XPath 递归搜索** |

### 待决事项

- [ ] ReviewGate 的 API 模式交互协议（proposal_id 生命周期）
- [ ] 表格布局的数据格式定义
- [ ] 代码块布局的语法高亮方案
- [ ] 是否支持从模板提取动画预设
- [ ] 是否支持多 LOGO（如公司 LOGO + 产品 LOGO）
- [ ] TextMeasurer 是否用 Pillow render 精确测量替代经验值估算
- [ ] Connector 避让算法：简单中间点 vs A* 寻路
- [x] 图形是否支持 add_group 打组（已确认：需先 add shapes 到 slide 再打组）
- [x] 循环图的弧形箭头实现：build_freeform + add_line_segments 多段折线近似（无原生 Bezier）
- [x] 分页拆分后 EnterpriseDesignDecider 如何同步生成额外 PageDesign（见 §12.13 集成逻辑）
- [ ] 跨页连接的视觉提示样式（"→ 续下页" / "← 接上页"）是否由模板控制
- [x] `--content` 与项目文件夹中 `content.json` 同时存在时的解析优先级（`--content` 优先，见 §12.11）
- [x] Connector `begin_connect(shape, site_index)` 的 connection site 索引方向映射（ROUNDED_RECTANGLE: 0=顶,1=左,2=底,3=右，见 §12.10）
- [ ] 教学动画的 XML 模板验证（需确认 PPT 打开后的实际效果）
- [x] `--history` 模式下 `query` 参数是否可选（nargs='?' + 验证，见 §7.4）
- [x] theme part 获取方式（已验证：prs.part.part_related_by(RT.THEME) + etree.fromstring(blob)，见 §3.2）
- [x] a:clrScheme/a:fontScheme 在 theme XML 中的位置（a:themeElements 子元素，sysClr 需 fallback，见 §3.2）
- [x] LOGO 检测规则（增加 PICTURE placeholder + layout-level 检测，见 §3.7）
- [x] PageContent 扩展字段 + EnterprisePageDesign 子类（见 §12.11）
- [x] content.json slides[] 解析实现方案（_load_enterprise_content()，见 §12.11）
- [x] 页面级增删改支持（--pages 参数 + PageRevisionEngine，见 §7.5 和 §9.6）
- [x] 风格一致性保证（BrandSpec + layout_id + density 三层机制，见 §9.6.6）
- [x] slide 删除/插入/替换 API 验证（全部通过，见 §9.6.3）
- [ ] 页面级修订时 ContentGenerator 的 "修改模式" prompt 设计（如何在旧内容基础上修改而非从零生成）
- [ ] INSERT 操作的 layout 自动选择策略（如何与相邻页面风格协调）

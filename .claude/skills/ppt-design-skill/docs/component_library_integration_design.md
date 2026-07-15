# Component Library Integration Design

## 1. 目标

将GroupShape/SmartArt组件数据库集成到PPT制作全流程，在opencode/claude code等skill环境下，让AI agent能：

1. **查询**：了解组件库有哪些可视化资源
2. **选用**：在制作PPT时主动选择合适的组件替代纯文字
3. **升级**：在beautify模式下，用组件库升级粗糙内容

## 2. 系统架构认知

当前系统有两种使用方式，本质相同——都是纯代码执行，不内置LLM：

| 入口 | 方式 | 示例 |
|------|------|------|
| CLI | `python -m ppt_pro_max "AI pitch"` | 命令行直接执行 |
| API | `generate_ppt("AI pitch")` | Python代码调用 |
| Skill | Agent(opencode)通过API调用 | 与API相同，Agent做决策 |

系统的4阶段Pipeline全部基于规则/模板，不调用LLM：

| 阶段 | 实现 | 说明 |
|------|------|------|
| StoryPlanner | 策略模板匹配 | 基于query关键词选择预定义策略（如YC Seed Deck） |
| DesignDecider | 设计规则 | 基于goal/emotion选择布局和配色 |
| ContentGenerator | 文案模板 | 基于goal/context拼模板文案 |
| PPTRenderer | python-pptx渲染 | 直接渲染 |

**LLM仅用于图片生成**（Seedream/GPT Image/DALL-E/Wanx/Kimi），需用户配置API Key。

因此组件库集成的核心是：
1. **暴露API让Agent/CLI查询资源** — catalog/search
2. **content.json支持组件字段** — 谁提供content（Agent或模板）谁决定用什么组件
3. **PrecisionRenderer执行组件注入** — 匹配→填数据→注入，纯执行逻辑

## 3. 当前状态

### 已完成（P9-P14）

| 模块 | 文件 | 说明 |
|------|------|------|
| GroupExtractor | `enterprise/group_extractor.py` | 提取完整`<p:grpSp>` XML + 图片blobs |
| SmartArtExtractor | `enterprise/smartart_extractor.py` | 提取4 XML parts |
| ComponentLibrary | `enterprise/component_library.py` | SQLite索引 + gzip存储 + 批量导入 + 去重 |
| ComponentRenderer | `enterprise/component_renderer.py` | 匹配→填文本→品牌色→注入 |
| build_library | `scripts/build_library.py` | 断点续传+后台运行 |

### 已验证

- 提取→入库→匹配→填数据→注入→保存→重开 ✅
- gzip压缩6.5倍，5文件导入5.6s ✅
- 600 tests passed, lint clean ✅

### 未集成（断路）— 已全部修复 ✅

| 断点 | 说明 | 修复 |
|------|------|------|
| Agent无法查询组件库 | 没有暴露catalog/search的公开API | ✅ `query_component_library()` + `catalog()` + CLI flags |
| content.json不支持组件 | page格式没有component_type字段 | ✅ 3个component字段透传 |
| PrecisionRenderer不查库 | render_slide()只画文字，从不查组件库 | ✅ `component_lib`参数 + 优先级链 |
| Pipeline不初始化库 | generate_ppt()不传递ComponentLibrary | ✅ `find_db_path()`自动查找 + 传递 |

## 4. 集成设计

### 4.1 公开API（新增）

CLI和API/Agent均可使用：

```bash
# CLI查询
python -m ppt_pro_max --component-catalog
python -m ppt_pro_max --component-search-type group --component-search-category process --component-search-nodes 4
```

```python
# API/Agent查询
from ppt_pro_max import generate_ppt, query_component_library

# 1. 查询组件库资源概览
catalog = query_component_library()
# → {
#     "group": {
#       "process": {"count": 144, "node_counts": "2-99"},
#       "swot":    {"count": 25,  "node_counts": "2-65"},
#       ...
#     },
#     "smartart": { ... }
#   }

# 2. 搜索特定类型
results = query_component_library(type="group", category="process", node_count=4)
# → [{"id": 42, "category": "process", "node_count": 4, "variant": "chevron"}, ...]

# 3. 通过content指定组件（API/Agent直接传dict，无需写文件）
result = generate_ppt(
    "产品发布",
    content={
        "slides": [
            {
                "goal": "content",
                "title": "开发流程",
                "bullets": ["需求分析", "设计", "开发", "测试"],
                "component_type": "group",
                "component_category": "process"
            }
        ]
    }
)
```

**新增参数说明**：

| 参数 | 位置 | 类型 | 说明 |
|------|------|------|------|
| `content` | `generate_ppt()` | `dict \| None` | 直接传入content dict，无需写JSON文件。与`content_file`互斥，`content`优先 |
| `--component-catalog` | CLI | flag | 打印组件库资源概览后退出 |
| `--component-search-type` | CLI | `str` | 搜索过滤：组件类型（group/smartart） |
| `--component-search-category` | CLI | `str` | 搜索过滤：分类（process/swot/...） |
| `--component-search-nodes` | CLI | `int` | 搜索过滤：节点数 |

**db_path查找公共函数**（`component_library.py`新增）：

```python
def find_db_path(component_library: str | None = None, project_dir: str | None = None) -> str | None:
    """三级查找组件库db路径，供query_component_library()和generate_ppt()共用"""
    # 1. 用户显式传入
    if component_library and os.path.exists(component_library):
        return component_library
    # 2. 项目目录下
    if project_dir:
        p = os.path.join(project_dir, "component_library", "index.db")
        if os.path.exists(p):
            return p
    # 3. 包目录下
    pkg_dir = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(pkg_dir, "component_library", "index.db")
    if os.path.exists(p):
        return p
    # 4. 未找到
    return None
```

### 4.2 content.json扩展

page新增3个可选字段：

```json
{
  "goal": "content",
  "title": "开发流程",
  "bullets": ["需求分析", "设计", "开发", "测试"],
  "component_type": "group",
  "component_category": "process",
  "component_variant": "chevron"
}
```

| 字段 | 必填 | 说明 |
|------|------|------|
| component_type | 否 | "group"或"smartart"，空则由renderer自动判断 |
| component_category | 否 | process/swot/hierarchy/infographic/chart/timeline |
| component_variant | 否 | 布局变体，空则自动选择 |

### 4.3 PrecisionRenderer集成

render_slide()增加component_lib参数，**纯执行，不做决策**：

```
render_slide(prs, page, component_lib=None)
  │
  ├─ is_hero (hook/cta)?
  │    → 走hero路径（不变）
  │
  └─ content路径，优先级链：
       cards > component_type > diagram_type > code > exercise > bullets
       │            │                │
       │            ▼                ▼
       │    ComponentRenderer    DiagramEngine
       │    (组件库XML注入)     (从零绘制简单图形)
       │
       └─ component_type分支：
            → ComponentRenderer.render(slide, element, brand_spec, lib)
            → lib.match({type, category, node_count=len(bullets)})
            → 匹配成功 → 填文本 → 品牌色 → 注入
            → 匹配失败 → 降级尝试diagram_type → DiagramEngine
            → diagram_type也无 → fallback画bullets文字
```

**component_type与diagram_type协作规则**：

| page字段 | 行为 |
|----------|------|
| 有component_type，匹配成功 | 用组件库GroupShape/SmartArt（最高质量） |
| 有component_type，匹配失败，有diagram_type | 降级到DiagramEngine（中等质量） |
| 有component_type，匹配失败，无diagram_type | fallback画bullets文字（兜底） |
| 无component_type，有diagram_type | 走原有DiagramEngine逻辑（不变） |
| 都无 | 走原有bullets/cards/code/exercise逻辑（不变） |

**bounds计算**：组件注入需要位置和尺寸，由render_slide()根据goal类型计算，复用现有布局区域划分逻辑：

```python
# render_slide()中component_type分支的bounds计算
if page.get("component_type"):
    content_area = self._get_content_area(slide, goal)  # 复用现有区域计算
    bounds = content_area  # (x, y, w, h) in inches
    element = {
        "type": page["component_type"],
        "category": page.get("component_category", ""),
        "variant": page.get("component_variant", ""),
        "texts": page.get("bullets", []),
        "bounds": bounds,
    }
    success = ComponentRenderer().render(slide, element, self.brand_spec, component_lib)
    if not success and page.get("diagram_type"):
        # 降级到DiagramEngine
        ...
```

注意：**不自动推断category**。Agent没指定component_type就说明Agent认为用文字更合适，Pipeline尊重Agent的决策。但beautify full模式例外——由规则引擎推断（见4.5）。

### 4.4 generate_ppt() API扩展

```python
# 新增参数
def generate_ppt(
    query, ...,
    component_library: str | None = None,  # 组件库db路径，None=自动查找
):
```

db_path查找顺序：
1. 用户显式传入 `component_library` 参数
2. 项目目录下 `component_library/index.db`
3. 包目录下 `component_library/index.db`
4. 不使用组件库（component_lib=None）

### 4.5 Beautify模式重新定义

**核心原则**：保留企业VI/LOGO，用组件库升级粗糙内容。

| 保留（不动） | 升级（重做） |
|-------------|-------------|
| 企业LOGO位置和图片 | 纯文字段落 → 图形化（组件库匹配） |
| VI配色 → 提取为BrandSpec | 简陋布局 → 专业布局 |
| 企业指定图片 | 增加配图/SmartArt |
| 母版/模板结构 | 增加设计元素 |

两种模式通过参数控制：

```python
# 轻量换皮：只改颜色字体
generate_ppt("美化", beautify="client.pptx", beautify_mode="light")

# 完整升级：内容梳理+图形化+配图
generate_ppt("美化", beautify="client.pptx", beautify_mode="full")

# 默认=full
generate_ppt("美化", beautify="client.pptx")
```

**full模式流程**：
1. 提取品牌元素（LOGO、VI颜色、字体）→ BrandSpec
2. 提取内容（SlideExtractor → page dicts，与content_parser产出相同格式）
3. 内容梳理（规则推断，与content_parser的goal推断逻辑一致）：
   - 推断goal：关键词匹配（同content_parser的中英文关键词表）
   - 推断component_type+component_category：基于bullets数量和结构特征
     - 3-8个顺序bullets → `component_type="group", component_category="process"`
     - 4项+对比/优势/劣势关键词 → `component_type="group", component_category="swot"`
     - 层级缩进结构 → `component_type="group", component_category="hierarchy"`
     - 数据表格 → `component_type="smartart", component_category="matrix"` 或保留chart
     - 其他 → 不设component_type，保持文字
4. 图形化（文字→ComponentLibrary匹配注入GroupShape/SmartArt）
5. 增加配图（ImageFetcher）
6. 用提取的BrandSpec统一风格
7. 保留LOGO和企业图片不动

**light模式流程**（当前inplace_restyle）：
1. 复制原始PPT
2. 替换schemeClr为BrandSpec颜色
3. 替换主题字体
4. 替换背景色
5. 保留所有shapes不变

**内容梳理的统一逻辑**：无论源素材是PPT（SlideExtractor）还是md/doc（ContentParser），产出都是page dicts，后续推断逻辑完全一致——都是基于规则的关键词匹配和结构特征判断，不依赖LLM。

### 4.6 数据流

```
[素材库PPT] ──build_library──→ [ComponentLibrary: SQLite+gzip]
                                      │
                    ┌─────────────────┼─────────────────┐
                    ▼                                     ▼
            CLI / API / Agent                      content.json
            query_component_library()               component_type字段
            → 看看有什么可用                      → 指定用什么组件
                    │                                     │
                    └─────────────────┬───────────────────┘
                                      ▼
                              generate_ppt()
                              → PrecisionRenderer
                              → page有component_type → 查库→注入
                              → page无component_type → 原有逻辑
                                      │
                                      ▼
                                 [输出PPT]
```

## 5. 实现步骤

### Step 1: ComponentLibrary.catalog() + find_db_path() ✅

文件：`enterprise/component_library.py`

- 新增`catalog()`方法：按type→category分组统计count和node_count范围，返回精简摘要，供agent快速了解资源。内存缓存，add/remove时失效
- 新增`find_db_path()`模块级函数：三级查找组件库db路径（用户指定→项目目录→包目录），供`query_component_library()`和`generate_ppt()`共用，避免查找逻辑重复

### Step 2: 公开API ✅

文件：`__init__.py` + `cli.py`

新增`query_component_library()`函数：
- 调用`ComponentLibrary.catalog()`
- 支持按type/category/node_count搜索
- 通过`find_db_path()`自动查找db_path

新增`content: dict | None`参数到`generate_ppt()`：
- 与`content_file`互斥，`content`优先
- 传入dict时写入临时JSON文件再走现有流程（最小改动），atexit自动清理
- Agent可直接传dict指定component_type，无需写文件

CLI新增参数：
- `--component-catalog`：打印概览后退出
- `--component-search-type` / `--component-search-category` / `--component-search-nodes`：搜索过滤

### Step 3: content.json格式扩展 ✅

文件：`enterprise/content_parser.py` + `enterprise/pipeline.py`

- `load_enterprise_content()`解析时识别component_type/component_category/component_variant，透传给render_slide()
- `_build_page_designs()`复制字段时增加3个component字段（当前不复制会被丢弃）
- 新增`infer_component_category()`函数：基于关键词+结构特征的规则推断

### Step 4: PrecisionRenderer集成 ✅

文件：`enterprise/precision_renderer.py`

- render_slide()新增component_lib参数
- content路径优先级链：`cards > component_type > diagram_type > code > exercise > bullets`
- page有component_type → ComponentRenderer注入，匹配失败降级到diagram_type → DiagramEngine → bullets
- bounds由render_slide()根据goal类型计算，复用现有布局区域划分逻辑
- **不做自动推断**，尊重Agent决策（beautify full模式例外，由规则引擎推断）

### Step 5: generate_ppt()集成 ✅

文件：`__init__.py`

- 新增component_library参数
- 通过`find_db_path()`自动查找db_path
- 初始化ComponentLibrary传递给PrecisionRenderer
- content dict参数处理：写入临时JSON → 走现有content_file流程，atexit自动清理

### Step 6: Beautify模式 ✅

文件：`enterprise/pipeline.py`

- 新增`beautify_mode`参数（"light"/"full"，默认"full"）
- 提取品牌元素 → BrandSpec（保留企业VI/LOGO）
- light模式：inplace_restyle换皮（原run_beautify逻辑，不变）
- full模式：
  1. SlideExtractor提取内容 → page dicts
  2. 规则推断component_type/component_category（与content_parser的goal推断逻辑一致，基于关键词+结构特征）
  3. PrecisionRenderer重建（component_lib注入组件）
  4. 提取失败时fallback到light模式
- run_beautify()拆分为_beautify_light() + _beautify_full()

### Step 7: 测试+验证 ✅

- 38个测试在test_component_integration.py
- catalog()缓存/失效测试（add/remove/bulk_import）
- find_db_path()三级查找测试
- query_component_library() API测试（catalog/search/空库）
- content_parser component字段透传测试
- _build_page_designs() component字段复制测试
- render_slide()优先级链测试（cards>component_type>diagram_type>bullets）
- render_slide()降级测试（匹配失败→diagram→bullets）
- generate_ppt()新参数测试（component_library/content dict）
- beautify_mode测试（light/full/默认/API）
- infer_component_category()推断测试（process/swot/hierarchy/边界）
- 边界测试（component_type=None, 无component_lib, 错误路径）
- 全量638 tests passed, lint clean

## 6. 性能数据

| 指标 | 当前值 |
|------|--------|
| 124个PPT GroupShape总数 | 17,337个 |
| 入库（去重，node_count>=1） | 预估5,000-8,000个 |
| 入库速度 | 1.1s/file（gzip+事务） |
| gzip压缩率 | 15.4%（6.5倍压缩） |
| 单组件注入时间 | <0.1s |
| catalog()查询 | <10ms |

## 7. 风险与约束

1. **组件质量**：入库的GroupShape质量参差不齐
   - 缓解：后续可加评分/审核机制，catalog()只展示优质组件

2. **匹配失败兜底**：Agent指定了component_type但库中无匹配
   - 缓解：PrecisionRenderer匹配失败时fallback画文字，不报错

3. **图片引用**：GroupShape中的图片rId需要重映射
   - 已解决：ComponentRenderer._inject_group_to_slide()处理了rId重映射

4. **文本替换精度**：按顺序替换`<a:t>`可能错位
   - 缓解：后续可基于level/role精确匹配

5. **db_path稳定性**：ComponentLibrary需要稳定的db路径
   - 缓解：三级查找（用户指定→项目目录→包目录）

6. **skill环境约束**：opencode/claude code下agent的调用方式
   - 通过python API调用，与CLI等价
   - Pipeline不做LLM决策，系统没有内置LLM
   - LLM仅用于图片生成，需用户配置API Key

7. **component_type与diagram_type冲突**：page同时指定两者时的行为
   - 规则：component_type优先，匹配失败降级到diagram_type，再失败fallback画文字
   - 组件库是DiagramEngine的升级替代，两者不冲突，形成质量梯度

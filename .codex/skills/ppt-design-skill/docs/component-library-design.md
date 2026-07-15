# 组件库设计方案

> 独立于修订方案（`revision-plan.md`）的组件库详细设计文档。
> 修订方案中 8.7 节的验证结论引用本文档。

---

## 一、目标

管理数万个 SmartArt / GroupShape / OLE 模板组件，支持：
- **高效提取**：从客户 PPT 中提取 SmartArt/GroupShape 的文本、结构、样式
- **高效检索**：按类型/分类/节点数/标签从万级组件库中毫秒级查找
- **高效加载**：按需加载 XML 模板，不预加载全部
- **高效编辑**：填入数据、应用品牌色/字体后写入新 PPT

---

## 二、技术可行性验证

### 2.1 验证方法

构造含 SmartArt 和 GroupShape 的 .pptx 文件，分别用 python-pptx 高层 API 和 lxml 低层 XML 解析测试。

测试脚本：
- `tests/test_smartart_inject.py` — SmartArt XML 注入 + 提取验证
- `tests/test_groupshape_extract.py` — GroupShape XML 注入 + 提取验证

### 2.2 验证结果

| 元素类型 | python-pptx 高层 API | lxml 低层 XML 解析 | 验证结果 |
|---------|---------------------|-------------------|---------|
| SmartArt 文本 | **不可见**（0 shapes） | `ppt/diagrams/dataN.xml` → `dgm:pt/a:t` | ✓ 4/4 文本提取成功 |
| SmartArt 类型 | **不可见** | `ppt/diagrams/layoutN.xml` → `uniqueId` + `dgm:alg/@type` | ✓ 识别出 lin（线性流程） |
| GroupShape 文本 | `shape.shapes` 可遍历 | `shape._element.findall(".//a:t")` | ✓ 3/3 文本提取成功 |
| GroupShape 结构 | `shape.shapes` 返回子列表 | XML 层级遍历 grpSp/sp | ✓ 递归提取成功 |

### 2.3 Round-Trip 验证（提取→修改→注入→PowerPoint 渲染）

测试脚本：`tests/test_smartart_roundtrip.py`

从 art.pptx 提取 3 个 SmartArt → 修改配色（accent1→D97706 琥珀色）+ 文本 → 注入新 PPT → PowerPoint 打开验证。

| 操作 | 修改目标 | 关键发现 | 结果 |
|------|---------|---------|------|
| 文本替换 | `dataN.xml` 的 `dgm:pt/a:t` | pt 的 type 属性为空（非 "node"），需按"有 a:t 且非 doc/parTrans/sibTrans/pres"筛选 | ✓ |
| 文本替换 | `drawingN.xml` 的 `dsp:txBody/a:t` | **drawing 含预渲染文本副本，PowerPoint 显示的是 drawing 里的文本** | ✓ |
| 配色替换 | `colorsN.xml` 的 `schemeClr` | **PowerPoint 从 colors XML 重新计算渲染**。必须将 schemeClr 替换为 srgbClr | ✓ 131→0 schemeClr |
| 注入新 PPT | slide XML + rels + Content_Types | diagram MIME 类型必须精确匹配 | ✓ PowerPoint 正确渲染 |

**测试 4：drawing.xml 必要性验证（test_smartart_no_drawing.py）**

| 测试 | drawing.xml | colors.xml | 结果 |
|------|------------|-----------|------|
| A | **无** | 修改原始（schemeClr→srgbClr） | ✓ 正确渲染，配色为琥珀色 |
| B | **无** | 从零生成（简化版 39 个颜色） | ✗ 全黑（styleLbl 不完整） |
| C | 有 | 从零生成（简化版 39 个颜色） | ✗ 全黑（与 B 完全相同，drawing 无影响） |

结论：**drawing.xml 不需要存储/注入**（PowerPoint 自动重建）；**colors.xml 不能从零生成**（必须存原始版）。

**Round-Trip 关键发现**：

1. **配色必须改 colorsN.xml**：PowerPoint 渲染 SmartArt 时从 colors XML 读取 schemeClr 再映射主题色。正确做法：克隆原始 colors XML → 将 `<a:schemeClr val="accent1"/>` 替换为 `<a:srgbClr val="D97706"/>`
2. **drawing.xml 不需要**：PowerPoint 从 layout+data+colors+quickStyle 自动重建渲染。不注入 drawing 时文本只需改 data 即可
3. **colors.xml 必须存原始版**：原始文件有 131 个颜色映射覆盖所有 styleLbl，从零生成的简化版（39 个）会导致全黑
4. **文本只需改 data**：不注入 drawing 时，修改 dataN.xml 的文本后 PowerPoint 自动同步到渲染
5. **data XML 的 pt 节点 type 不是 "node"**：实际文本节点 type 为空字符串，需按"有 a:t 且 type ∉ {doc, parTrans, sibTrans, pres}"筛选
6. **4 种 diagram MIME 类型**：data→diagramData, layout→diagramLayout, quickStyle→diagramStyle, colors→diagramColors
7. **schemeClr 角色映射**：colors XML 中出现的角色有 accent1-6、lt1、dk1、tx1，品牌配色映射需覆盖全部

### 2.4 .pptx 内部结构

.pptx 是 ZIP 包，SmartArt 相关文件（PowerPoint 实际生成的目录名是 `ppt/diagrams/`）：

```
ppt/slideN.xml          → 引用 SmartArt 的 rId
ppt/slideN.xml.rels     → rId → "../diagrams/dataN.xml" + layoutN.xml + ...
ppt/diagrams/dataN.xml       → 逻辑数据（dgm:pt / a:t 节点）
ppt/diagrams/layoutN.xml     → 布局模板（uniqueId + dgm:alg type=lin/cycle/hierarchy/...）
ppt/diagrams/colorsN.xml     → 配色方案（schemeClr 角色，注入时克隆替换）
ppt/diagrams/quickStyleN.xml → 快速样式
ppt/diagrams/drawingN.xml    → 预渲染缓存（PowerPoint 自动重建，不需要存储/注入）
```

GroupShape 在 slide XML 内：

```xml
<p:grpSp>                          <!-- GroupShape 容器 -->
  <p:nvGrpSpPr>...</p:nvGrpSpPr>   <!-- 名称/ID -->
  <p:grpSpPr>                      <!-- 变换（位置/尺寸/子坐标系） -->
    <a:xfrm>...</a:xfrm>
  </p:grpSpPr>
  <p:sp>...</p:sp>                 <!-- 子形状1（文本框/矩形等） -->
  <p:sp>...</p:sp>                 <!-- 子形状2 -->
  <p:grpSp>...</p:grpSp>           <!-- 嵌套子组（递归） -->
</p:grpSp>
```

---

## 三、三层架构

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

---

## 四、XML 提取层

### 4.1 SmartArtExtractor

```python
class SmartArtExtractor:
    """从 .pptx 的 ZIP 内 XML 提取 SmartArt 文本和结构。"""

    def extract(self, pptx_path: str, slide_index: int, shape_rId: str) -> dict:
        """
        流程：
        1. zipfile 打开 .pptx
        2. 从 slideN.xml.rels 找到 diagrams/ 下的 5 个 XML 文件
        3. 解析 dataN.xml → 提取所有文本节点（dgm:pt / a:t）
           注意：文本节点的 type 属性为空字符串，不是 "node"
           需按"有 a:t 子元素且 type 不在 {doc, parTrans, sibTrans, pres} 中"筛选
        4. 解析 layoutN.xml → 提取 uniqueId（精确类型名）+ alg/@type（大类）
        5. 解析 colorsN.xml → 提取 schemeClr 角色映射（accent1-6, lt1, dk1, tx1）
           注意：colors.xml 必须存原始版（131 个颜色映射），不能从零生成
        6. 返回结构化数据（4 个 XML parts，不需要 drawing）
        """

    def extract_all(self, pptx_path: str) -> list[dict]:
        """扫描所有 slide，提取全部 SmartArt。"""

    def _parse_data(self, data_xml: bytes) -> list[dict]:
        """解析 dataN.xml，返回 [{"id": 0, "text": "...", "level": 0}, ...]"""

    def _parse_layout(self, layout_xml: bytes) -> dict:
        """解析 layoutN.xml，返回 {"category": "process", "variant": "chevron", "node_count": 4}"""

    def _parse_colors(self, colors_xml: bytes) -> dict:
        """解析 colorsN.xml，返回配色角色映射 {"fill": "accent1", "line": "lt1", "text": "tx1"}"""
```

**SmartArt 类别推断**（从 layout XML 的 `dgm:alg/@type`）：

| alg type 值 | 推断类别 | DiagramEngine 对应 |
|-------------|---------|-------------------|
| lin | process（线性流程） | process |
| snake | process（蛇形流程） | process |
| cycle | cycle（循环） | cycle |
| hier | hierarchy（层级/组织架构） | tree |
| pyra | pyramid（金字塔） | pyramid |
| matrix | matrix（矩阵） | swot |
| seg | relationship（关系） | — |
| pic | picture（图片布局） | — |

### 4.2 GroupExtractor

```python
class GroupExtractor:
    """递归遍历 GroupShape 内部元素。"""

    def extract(self, group_shape) -> dict:
        """
        递归遍历 GroupShape：
        - 子文本框 → texts[]
        - 子图片 → images[]（保存 blob 到临时文件）
        - 子形状（矩形/箭头/连接线）→ shapes[]（类型+位置）
        - 嵌套 GroupShape → 递归
        """

    def _traverse_group_element(self, grp_elem, ns: dict) -> dict:
        """从 lxml 元素遍历，不依赖 python-pptx。"""

    def _infer_category(self, structure: dict) -> str:
        """从形状排列推断类别：
        - 多个矩形横向排列 + 箭头 → process
        - 多个矩形纵向排列 → hierarchy
        - 同心圆/环形 → cycle
        - 嵌套矩形 → matrix
        """
```

### 4.3 OLEExtractor

```python
class OLEExtractor:
    """提取嵌入对象元信息。"""

    def extract(self, pptx_path: str, slide_index: int, shape_rId: str) -> dict:
        """
        1. 从 slideN.xml.rels 找到 embedded 对象
        2. 读取 oleObject 的 progId（如 "Excel.Sheet.12"）
        3. 尝试读取关联的 chart XML（如果有）
        4. 返回元信息
        """
```

---

## 五、组件库层

### 5.1 存储方案

SQLite 索引 + 文件系统存 XML：

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

### 5.2 SQLite Schema

```sql
CREATE TABLE components (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    type        TEXT NOT NULL,           -- 'smartart' / 'group' / 'ole'
    category    TEXT NOT NULL,           -- 'process' / 'hierarchy' / 'cycle' / ...
    variant     TEXT,                    -- 'chevron' / 'arrows' / 'orgchart' / ...
    node_count  INTEGER,                 -- 节点数量
    level_count INTEGER,                 -- 层级数量
    tags        TEXT,                    -- JSON array: ["flow","horizontal","arrow"]
    xml_path    TEXT NOT NULL,           -- SmartArt: 目录路径（含 4 个 XML parts）；Group/OLE: 单文件路径
    thumb_path  TEXT,                    -- 缩略图路径
    source      TEXT DEFAULT 'builtin',  -- 'builtin' / 'microsoft' / 'custom'
    checksum    TEXT,                    -- XML 内容 MD5（去重用）
    created_at  TEXT DEFAULT (datetime('now')),
    updated_at  TEXT DEFAULT (datetime('now'))
);

CREATE VIRTUAL TABLE components_fts USING fts5(
    id, type, category, variant, tags
);

CREATE INDEX idx_type_category ON components(type, category);
CREATE INDEX idx_node_count ON components(node_count);
CREATE INDEX idx_type_category_nodes ON components(type, category, node_count);
CREATE INDEX idx_checksum ON components(checksum);
```

### 5.3 ComponentLibrary API

```python
class ComponentLibrary:
    def __init__(self, db_path: str = "component_library/index.db"):
        self._db = sqlite3.connect(db_path)

    # ── 查询 ──

    def search(self, type: str, category: str, node_count: int = None,
               tags: list[str] = None, limit: int = 10) -> list[Component]:
        """按条件检索组件"""

    def match(self, extracted: dict) -> Component | None:
        """
        匹配引擎：给定提取的结构化数据，找最相似的组件
        1. 按 type + category 过滤
        2. 按 node_count 过滤
        3. 按变体名模糊匹配
        4. 返回最佳匹配或 None
        """

    def get(self, component_id: int) -> Component | None:
        """按 ID 获取"""

    # ── 加载 ──

    def load_xml(self, component_id: int) -> dict:
        """加载组件的 XML 模板（4 个 parts: data/layout/colors/quickStyle，不需要 drawing）"""

    def load_thumbnail(self, component_id: int) -> bytes | None:
        """加载缩略图"""

    # ── 写入 ──

    def add(self, type: str, category: str, variant: str,
            xml_parts: dict, thumb: bytes = None, **meta) -> int:
        """新增组件（SmartArt: xml_parts 含 4 个 XML parts；Group/OLE: 单 XML）"""

    def bulk_import(self, pptx_paths: list[str]) -> dict:
        """
        批量导入：从多个 PPT 中提取所有 SmartArt/Group → 入库
        流程：遍历 PPT → SmartArtExtractor.extract_all() → 去重(checksum) → 入库
        返回：{"added": 1523, "skipped": 47, "errors": 12}
        """

    def remove(self, component_id: int) -> bool:
        """删除组件"""

    # ── 统计 ──

    def stats(self) -> dict:
        """{"smartart": 8500, "group": 3200, "ole": 450, "total": 12150}"""

    def coverage(self, category: str) -> dict:
        """某类别的节点数覆盖情况：{"2": 15, "3": 22, "4": 35, ...}"""
```

### 5.4 性能设计

| 操作 | 数据量 | 目标延迟 | 实现方式 |
|------|--------|---------|---------|
| search(type, category, node_count) | 万级 | <5ms | SQLite 索引查询 |
| match(extracted) | 万级 | <10ms | 先按 category+node_count 缩小到百级，再排序 |
| load_xml(id) | 单个 | <3ms | 文件系统读取 4 个 XML parts（每个 <50KB） |
| bulk_import(paths) | 百个 PPT | 批量 | 逐 PPT 解析 + 事务批量 INSERT |
| add() | 单个 | <1ms | INSERT + 事务 |

### 5.5 版本管理

```sql
CREATE TABLE schema_version (
    version INTEGER PRIMARY KEY,
    applied_at TEXT DEFAULT (datetime('now')),
    description TEXT
);
```

```
component_library/
├── index.db
└── migrations/
    ├── v1_initial.sql
    ├── v2_add_level_count.sql
    └── ...
```

---

## 六、渲染桥接层

### 6.1 ComponentRenderer

```python
class ComponentRenderer:
    """提取 → 匹配组件库 → 填入数据 → 应用品牌 → 写入 slide"""

    def render(self, slide, complex_element: dict, brand_spec: BrandSpec,
               component_lib: ComponentLibrary) -> bool:
        """
        1. 从 complex_element 获取 type/category/nodes
        2. component_lib.match() 找到最匹配的组件模板
        3. 加载组件 XML（4 个 parts: data/layout/colors/quickStyle，不需要 drawing）
        4. 将 texts[] 填入 data XML 的文本节点（PowerPoint 自动同步到渲染）
        5. 应用 brand_spec 的配色到 colors XML（克隆原始→schemeClr → srgbClr）
        6. 将处理后的 XML 写入 slide
        """

    def render_smartart(self, slide, element: dict, brand_spec, lib) -> bool:
        """SmartArt 专用渲染路径"""
        component = lib.match(element)
        if component is None:
            return self._fallback_diagram(slide, element, brand_spec)

        xml_parts = lib.load_xml(component.id)  # {data, layout, colors, quickStyle}
        filled_parts = self._fill_data(xml_parts, element["nodes"])
        styled_parts = self._apply_brand_colors(filled_parts, brand_spec)
        self._inject_to_slide(slide, styled_parts, element["bounds"])
        return True

    def render_group(self, slide, element: dict, brand_spec, lib) -> bool:
        """GroupShape 渲染路径"""
        # GroupShape 通常用 PrecisionRenderer 的 add_card/add_rect 重建
        # 但也可以匹配组件库中的 Group 模板
        ...

    def _fallback_diagram(self, slide, element, brand_spec) -> bool:
        """
        组件库无匹配时，降级用 DiagramEngine 渲染：
        SmartArt process → DiagramEngine.render("process", data)
        SmartArt hierarchy → DiagramEngine.render("tree", data)
        SmartArt cycle → DiagramEngine.render("cycle", data)
        SmartArt matrix → DiagramEngine.render("swot", data)
        SmartArt pyramid → DiagramEngine.render("pyramid", data)
        SmartArt picture → 用 image + text 重建布局
        """

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
```

### 6.2 降级路径优先级

```
组件库匹配 → 用组件库模板重建（配色/字体统一）
    ↓ 无匹配
DiagramEngine 重建（结构近似，样式不同但可控）
    ↓ 不支持的 diagram_type
文本降级（提取的 texts 作为 bullets 显示）
```

---

## 七、批量导入流程

### 7.1 从现有 PPT 文件批量入库

```bash
# 命令行
python -m ppt_pro_max.component_importer --source ./microsoft_templates/ --library ./component_library/

# API
from ppt_pro_max.enterprise.component_library import ComponentLibrary
lib = ComponentLibrary("./component_library/index.db")
result = lib.bulk_import(["templates/microsoft_smartart.pptx"])
# → {"added": 1523, "skipped": 47, "errors": 12}
```

### 7.2 从客户 PPT 提取新组件

美化改版时，遇到组件库没有的 SmartArt/Group，自动提示入库：

```python
extractor = SmartArtExtractor()
element = extractor.extract("client.pptx", slide_idx=3, shape_rId="rId3")
if lib.match(element) is None:
    # 自动入库（可选，需用户确认）
    lib.add(type="smartart", category=element["category"],
            variant=element["variant"], xml_parts=element["xml_parts"])
```

### 7.3 缩略图生成

| 方案 | 优点 | 缺点 |
|------|------|------|
| comtypes/win32com → 导出图片 | 精确渲染 | Windows only，需 PowerPoint |
| LibreOffice headless → 导出图片 | 跨平台 | 需安装 LibreOffice |
| Pillow 画简易示意图 | 最快，无依赖 | 不精确，仅示意 |

建议：开发期用方案 3（Pillow），生产环境用方案 1 或 2。

---

## 八、新增文件清单

| 文件 | 说明 |
|------|------|
| `src/ppt_pro_max/enterprise/smartart_extractor.py` | SmartArt XML 解析提取 |
| `src/ppt_pro_max/enterprise/group_extractor.py` | GroupShape 递归提取 |
| `src/ppt_pro_max/enterprise/ole_extractor.py` | OLE/嵌入对象提取 |
| `src/ppt_pro_max/enterprise/component_library.py` | 组件库：SQLite 索引 + CRUD + 匹配 |
| `src/ppt_pro_max/enterprise/component_renderer.py` | 组件渲染桥接 |
| `src/ppt_pro_max/enterprise/component_importer.py` | 批量导入工具 |
| `component_library/` | 组件库存储目录 |
| `component_library/migrations/` | Schema 迁移脚本 |

---

## 九、实施顺序

| 阶段 | 内容 | 依赖 | 预估代码量 |
|------|------|------|-----------|
| P12 | SmartArt/Group XML 提取器 | P9 | ~250 行 |
| P13 | 组件库基础设施（Schema + Index + CRUD） | P12 | ~300 行 |
| P14 | 组件库匹配引擎 + 渲染桥接 | P13 | ~200 行 |
| P15 | 批量导入工具 + CLI | P13 | ~150 行 |
| P16 | 缩略图生成 + 组件库 UI 预览 | P14 | ~100 行 |

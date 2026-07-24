# Findings — 多画布比例支持审计

## 硬编码坐标审计结果

### 按文件统计

| 文件 | PROPORTIONAL | LAYOUT-STRATEGY | FONT-SIZE | 合计 |
|------|-------------|-----------------|-----------|------|
| layout_registry.py | 80 | 20 | 20 | 120 |
| precision_renderer.py | 18 | 12 | 0 | 30 |
| block_renderer.py | 6 | 10 | 0 | 16 |
| enterprise_renderer.py | 12 | 3 | 0 | 15 |
| design_dna_extractor.py | 10 | 5 | 0 | 15 |
| layout_engine.py | 12 | 2 | 0 | 14 |
| component_renderer.py | 5 | 2 | 0 | 7 |
| build_helpers.py | 5 | 1 | 0 | 6 |
| ppt_renderer.py | 5 | 0 | 0 | 5 |
| theme_composer.py | 3 | 3 | 0 | 6 |
| chart_builder.py | 4 | 0 | 0 | 4 |
| decoration_renderer.py | 1 | 0 | 0 | 1 |
| group_extractor.py | 1 | 0 | 0 | 1 |
| page_revision.py | 3 | 1 | 0 | 4 |
| slide_extractor.py | 1 | 0 | 0 | 1 |
| density_profile.py | 0 | 0 | 8 | 8 |
| **合计** | **166** | **59** | **28** | **253** |

### 关键硬编码值频率

| 值 | 含义 | 出现次数 |
|----|------|---------|
| 0.9 | 左/右边距 | ~25 |
| 11.533 | 内容宽度 (13.333-0.9-0.9) | ~20 |
| 1.6 | 内容区 Y 起始 | ~15 |
| 7.5 | 画布高度 | ~12 |
| 13.333 | 画布宽度 | ~10 |
| 7.0 | 页脚 Y (7.5-0.5) | 4 |
| 5.833 | 页脚居中 X (13.333/2-0.833) | 4 |
| 11.433 | 页脚右侧 X (13.333-0.9-1) | 4 |
| 5.5 | 半栏宽度 | 6 |
| 6.933 | 右栏 X (0.9+5.5+0.533) | 4 |
| 3.644 | 三栏卡片宽度 | 3 |
| 10.333 | 居中内容宽度 | 5 |
| 8.3 | 右侧图片 X | 3 |
| 4.2 | 右侧图片宽度 | 3 |

### LAYOUT-STRATEGY 类坐标详细清单

这些是需要布局策略引擎处理的坐标：

1. **双栏布局** (two-column, two-column-dense)
   - left_col: x=0.9, w=5.5
   - right_col: x=6.933, w=5.5
   - 竖屏→单栏堆叠

2. **三栏卡片** (three-column-cards)
   - card1: x=0.9, w=3.644
   - card2: x=4.844, w=3.644
   - card3: x=8.789, w=3.644
   - 竖屏→2列或1列

3. **四指标横排** (four-metrics)
   - metric1-4: x=0.9/4.108/7.317/10.525, w=2.708
   - 竖屏→2×2网格

4. **2×2网格** (grid-2x2-cards, swot-matrix)
   - card1-4: x=0.9/6.933, y=1.2/4.3, w=5.5, h=2.8
   - 竖屏→1列4行

5. **内容+侧图** (content-with-title)
   - body: x=0.9, w=7.0
   - side_visual: x=8.5, w=4.0
   - 竖屏→图上文下

6. **图表+洞察** (chart-focus)
   - chart: x=1.2, w=7.5
   - insight: x=9.2, w=3.5
   - 竖屏→图上洞察下

7. **图+文** (image-plus-text)
   - image: x=0, w=7.5
   - title/body: x=8.0, w=4.8
   - 竖屏→图上文下

8. **侧边栏** (sidebar-left)
   - sidebar: x=0, w=3.5
   - title/body: x=4.0, w=8.433
   - 竖屏→顶部栏

9. **右侧图片** (precision_renderer, pipeline)
   - image: x=8.3, w=4.2
   - 竖屏→图片在上方

10. **双栏项目符号** (precision_renderer)
    - col_w = (11.5 - 0.3) / 2
    - 竖屏→单栏

### API 透传路径

```
cli.py --aspect-ratio
  → generate_ppt(aspect_ratio=...)
    → SlideLayout.from_aspect_ratio(spec)
      → _generate_ppt_freestyle(slide_layout=...)
        → PPTRenderer(slide_layout=...)
        → PrecisionRenderer(slide_layout=...)
      → _generate_ppt_enterprise(slide_layout=...)
        → EnterprisePipeline.run(slide_layout=...)
          → PrecisionRenderer(slide_layout=...)
      → _generate_ppt_beautify(slide_layout=...)
        → EnterprisePipeline.run_beautify(slide_layout=...)
      → _generate_proposals(slide_layout=...)
        → ProposalGenerator.generate(slide_layout=...)
```

### Template 尺寸优先级

当 template.pptx 存在时：
- Template 自带尺寸 → 使用 Template 尺寸（不覆盖）
- aspect_ratio 参数 + Template → Template 优先，打印 warning
- aspect_ratio 参数 + 无 Template → 使用 aspect_ratio 指定尺寸
- 无 aspect_ratio + 无 Template → 默认 16:9

### 字体缩放公式

```python
def scale_font(self, size_pt: int) -> int:
    base_area = 13.333 * 7.5  # = 100
    current_area = self.width * self.height
    scale = (current_area / base_area) ** 0.5  # 面积的平方根
    return max(8, round(size_pt * scale))
```

各比例缩放效果：

| 比例 | 面积 | 缩放因子 | 44pt→ | 28pt→ | 16pt→ | 12pt→ |
|------|------|---------|-------|-------|-------|-------|
| 16:9 | 100 | 1.00 | 44 | 28 | 16 | 12 |
| 4:3 | 75 | 0.87 | 38 | 24 | 14 | 10 |
| 16:10 | 111 | 1.05 | 46 | 29 | 17 | 13 |
| 9:16 | 100 | 1.00 | 44 | 28 | 16 | 12 |
| 1:1 | 56.25 | 0.75 | 33 | 21 | 12 | 9 |

注意：9:16 面积与 16:9 相同（7.5×13.333 vs 13.333×7.5），所以字体缩放因子相同。
但 9:16 的宽度只有 7.5"，所以每行能容纳的字符更少，可能需要进一步减小字体。
→ 需要增加宽度感知的字体调整：当宽度 < 10" 时，额外乘以 width/13.333 的 0.3 次方。

修正公式：
```python
def scale_font(self, size_pt: int) -> int:
    base_area = 13.333 * 7.5
    area_scale = (self.width * self.height / base_area) ** 0.5
    width_scale = (self.width / 13.333) ** 0.3  # 宽度感知
    scale = area_scale * width_scale
    return max(8, round(size_pt * scale))
```

修正后效果：

| 比例 | 面积缩放 | 宽度缩放 | 综合缩放 | 44pt→ | 28pt→ | 16pt→ |
|------|---------|---------|---------|-------|-------|-------|
| 16:9 | 1.00 | 1.00 | 1.00 | 44 | 28 | 16 |
| 4:3 | 0.87 | 0.92 | 0.80 | 35 | 22 | 13 |
| 9:16 | 1.00 | 0.72 | 0.72 | 32 | 20 | 12 |
| 1:1 | 0.75 | 0.80 | 0.60 | 26 | 17 | 10 |

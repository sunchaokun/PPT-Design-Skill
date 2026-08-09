# 案例：CRISPR-Cas9 奠基论文 — 科研范式高密度 PPT

## 论文信息

| 项 | 内容 |
|----|------|
| **标题** | A Programmable Dual-RNA-Guided DNA Endonuclease in Adaptive Bacterial Immunity |
| **作者** | Martin Jinek, Krzysztof Chylinski, Ines Fonfara, Michael Hauer, Jennifer A. Doudna, Emmanuelle Charpentier |
| **期刊** | Science 337(6096): 816-821 (2012) |
| **DOI** | 10.1126/science.1225829 |
| **意义** | 首次证明双 RNA (crRNA + tracrRNA) 可编程引导 Cas9 进行序列特异性 DNA 切割，是 CRISPR 基因编辑技术的奠基工作；获 2020 年诺贝尔化学奖 |
| **方向** | 生物科技 / 分子生物学 / 基因编辑 |

## 核心科学发现

1. **双 RNA 引导**：Cas9 核酸内切酶需要 crRNA（提供靶向信息）与 tracrRNA（提供 Cas9 结合骨架）双 RNA 同时存在，才能切割靶 DNA
2. **PAM 识别**：靶位点下游存在 PAM 序列（化脓链球菌为 5'-NGG-3'），Cas9 通过识别 PAM 区分自身 DNA 与外源 DNA
3. **双结构域切割**：RuvC 结构域切割非靶标链，HNH 结构域切割靶标链，产生平末端；切割位点位于 PAM 上游 3 bp
4. **可编程性**：改变 crRNA 序列即可重定向 Cas9 到任意靶位点，奠定基因编辑技术基础

## PPT 设计范式（科研范式）

本案例严格遵循 **Scientific Research Paradigm**，与商业 PPT 完全不同：

| 维度 | 科研范式（本案例） | 商业范式（不使用） |
|------|-------------------|-------------------|
| 页面结构 | Figure + caption（期刊惯例） | KPI 卡片、Hero 页、特性卡片 |
| 颜色系统 | 语义生物学色：红=切割/上调、深蓝=对照、绿=引导、紫=突变 | 品牌强调色、渐变填充 |
| 引用 | 每个科学声明标注 (Author, Year) | 无引用 |
| 编号 | Figure 1-8、面板 A/B/C（科研必需） | "01/04" 商业卡片编号（科研禁用但商业也禁用） |
| 封面 | 论文标题 + 作者 + 机构（论文格式） | 营销 Hero 大图 |
| 动画 | 无（科研 PPT 必须可打印） | 入场/退场动画 |
| 数据展示 | 凝胶电泳、切割效率定量、序列比对、机制图 | 商业图表 |

## 10 页结构

| 页 | 标题 | 科研元素 | Figure |
|----|------|---------|--------|
| 1 | 封面（论文标题格式） | 标题 + 作者 + 机构 + 科学背景 | — |
| 2 | CRISPR 适应性免疫三阶段 | 机制流程图（适应/表达/干扰） | Figure 1 |
| 3 | 研究问题与假设 | 双问题卡片 + 核心假设 | Figure 2 |
| 4 | 体外重建方法 | 5 步实验流程图 | Figure 3 |
| 5 | 结果1：Cas9 依赖双 RNA | **凝胶电泳示意图**（6 泳道） | Figure 4 |
| 6 | 结果2：切割效率定量 | **原生柱状图**（双 RNA 95% vs 单 RNA 15%） | Figure 5 |
| 7 | 结果3：PAM 序列识别 | **序列比对表**（NGG/NAG/NGA 对比） | Figure 6 |
| 8 | 结果4：双结构域机制 | **机制图**（RuvC + HNH + 双链 DNA） | Figure 7 |
| 9 | 讨论与科学意义 | 4 象限意义分析 | Figure 8 |
| 10 | 结论 + 参考文献 | 4 结论 + 5 条参考文献 | Conclusion |

## 设计决策详解

### 语义颜色系统
```python
C = {
    'up_color': '#C0392B',      # 切割/上调 (红) — 表示 Cas9 切割活性
    'down_color': '#2C3E50',    # 对照/下调 (深蓝)
    'control_color': '#27AE60', # 引导/对照 (绿) — crRNA 相关
    'mutant_color': '#8E44AD',  # 突变体 (紫) — 单 RNA 条件
    'text_dark': '#2C3E50',     # 学术主文本
    'font_heading': 'Georgia',  # 学术衬线
    'font_cjk': '思源宋体',      # 中文衬线
    'font_mono': 'Consolas',    # 序列/代码
}
```
语义色传递生物学含义，而非装饰目的——这是科研 PPT 区别于商业 PPT 的核心。

### 凝胶电泳示意图（Figure 4）
6 个泳道用 `rect` 模拟琼脂糖凝胶：仅泳道 5（双 RNA + Cas9）出现两条切割产物条带，泳道 6（单 RNA）出现弱带，其余对照泳道为完整底物。条带位置模拟分子量差异。

### 序列比对表（Figure 6）
用等宽字体（Consolas）对齐 DNA 序列，通过改变 PAM 区域颜色突出 NGG/NAG/NGA 差异，右侧"切割"列用 ✓/✗ 标注活性。

### 机制图（Figure 7）
用色块表示 RuvC（红）/HNH（深蓝）结构域，双色条表示 DNA 双链（靶标链红、非靶标链深蓝），标注切割位点。

## BuildQA 验证

```python
from ppt_pro_max.build_qa import BuildQA

report = BuildQA().check('CRISPR-Cas9_Science_2012.pptx', mode='scientific')
print(report.is_passable)  # True, 0 fatal, 0 warning
```

**科研模式 (mode='scientific')**：BuildQA 将最小字号阈值从 11pt 调整为 8pt，因为期刊 Figure caption（9pt）、序列（8pt）、上标引用（8pt）是 Nature/Cell 标准排版。其余检查（占位文本、空白页、越界、图片拉伸）保持不变。

```python
# 商业 deck 用默认模式（11pt 阈值）
report = BuildQA().check('business_deck.pptx')          # mode='business' 默认
# 科研 deck 用科研模式（8pt 阈值）
report = BuildQA().check('paper_figures.pptx', mode='scientific')
```

该模式变更覆盖了 5 个专项测试（`test_build_qa_v2.py::TestScientificMode`）：9pt 在科研模式通过/商业模式报警、8pt 边界、7pt 仍报警（即使科研模式也有底线）、默认模式为 business。

## 复现

```bash
# 从项目根目录
pip install -e .
$env:PYTHONPATH = "src"
python showcase/crispr-cas9-2012/build_crispr.py

# 验证
python -c "
from ppt_pro_max.build_qa import BuildQA
r = BuildQA().check('showcase/crispr-cas9-2012/CRISPR-Cas9_Science_2012.pptx', mode='scientific')
print('PASS' if r.is_passable else 'FAIL')
"
```

## 输出文件

| 文件 | 说明 |
|------|------|
| `CRISPR-Cas9_Science_2012.pptx` | 最终 PPT（10 页，51.7 KB） |
| `build_crispr.py` | Build 模式源码（可编辑复现） |
| `README.md` | 本案例说明文档 |

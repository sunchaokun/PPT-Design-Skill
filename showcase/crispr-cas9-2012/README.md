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

## 设计原则：科研 PPT 是"多面板信息图表"，不是大图展示

本案例参考真实科研 deck（`E:\简历\output\案例_生物科学_CRISPR基因编辑_v4.pptx`）的版式，核心特征是：

**主流的科研 PPT 几乎不使用大面积图片**，而是用**多面板 (a/b/c/d) 信息图表**展示数据、模型、流程：

| 维度 | 科研范式实现 |
|------|-------------|
| 页面结构 | `Figure N` + 标题 + **a/b/c/d 多面板**（2×2 或左右分区） |
| 序列比对 | **逐碱基色块**（AUTO_SHAPE 方块）模拟 DNA/sgRNA 配对 |
| 数据展示 | 柱状图、折线图、对数坐标、大数字统计 |
| 流程模型 | 流程色块 + 箭头（RNP 组装、NHEJ/HDR 分叉） |
| 结构域 | 蛋白条带图（色块分段标注 REC/Bridge/RuvC/HNH/PI） |
| 底部 | `Note` + 引用（Jinek et al., Science 2012） |
| 图片 | **刻意不用大图**，全部用绘制元素表达信息 |

## 5 页结构（信息密度对标参考案例）

| 页 | Figure | 面板 | 信息密度 |
|----|--------|------|---------|
| 1 | 封面 | 论文标题 + 分子机制示意 | 75 shapes |
| 2 | Figure 1 | a 序列比对+切割  b 切割效率  c 双RNA组装  d 关键参数 | **164 shapes** |
| 3 | Figure 2 | a PAM变异扫描  b 切割位点  c Cas9结构域  d 关键统计 | **123 shapes** |
| 4 | Figure 3 | a 双结构域机制  b NHEJ/HDR修复  c 编辑效率  d 基因编辑应用 | 43 shapes |
| 5 | Figure 4 | a 错配敏感性折线  b 种子区效应  c 特异性指标  d 结论 | 39 shapes |

> 对比参考案例（25/171/128/139/94 shapes），核心结果页密度已达 90%+ 对标水平。

## 关键技术：BuildQA 修复

### 1. 科研模式 `mode='scientific'`
期刊 Figure caption（8-9pt）、序列（8pt）是标准排版，与商业 11pt 阈值冲突。新增科研模式（8pt 阈值），商业默认不变。

### 2. CJK 加权文本溢出检测
`_check_text_overflow` 原算法未考虑 CJK 字符全宽特性，导致中文 legend 被误报溢出。修复为：`CJK 字符权重 1.5× Latin`。

## 复现

```bash
# 从项目根目录
pip install -e .
python showcase/crispr-cas9-2012/build_crispr.py

# 验证 (科研模式)
python -c "
from ppt_pro_max.build_qa import BuildQA
r = BuildQA().check('showcase/crispr-cas9-2012/CRISPR-Cas9_Science_2012.pptx', mode='scientific')
print('PASS' if r.is_passable else 'FAIL')
"
```

## 输出文件

| 文件 | 说明 |
|------|------|
| `CRISPR-Cas9_Science_2012.pptx` | 最终 PPT（5 页，多面板信息图表） |
| `build_crispr.py` | Build 模式源码（可编辑复现） |
| `README.md` | 本案例说明文档 |

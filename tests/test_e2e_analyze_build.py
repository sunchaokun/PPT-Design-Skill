"""E2E test: analyze_template → build_helpers → output PPT"""
import os
import pytest
from pptx import Presentation


@pytest.fixture
def C():
    return {
        'primary': '#2E6504',
        'primary_mid': '#466740',
        'accent': '#7DA92F',
        'muted': '#84AF7D',
        'light': '#D4E3AC',
        'lighter': '#E7F3A6',
        'lightest': '#F2F8D6',
        'background': '#FFFFFF',
        'bg_tint': '#F2F8D6',
        'white': '#FFFFFF',
        'text_dark': '#0D4609',
        'text_body': '#2E6504',
        'text_muted': '#466740',
        'divider': '#84AF7D',
        'card_bg': '#F6FAE8',
        'highlight': '#7DA92F',
        'font_heading': 'Microsoft YaHei',
        'font_body': 'Microsoft YaHei',
    }


class TestE2EBuildFromAnalysis:

    def test_analyze_then_build(self, tmp_path, C):
        from ppt_pro_max.analyze_template import analyze
        from ppt_pro_max.build_helpers import (
            add_slide, rect, top_bar, page_header,
            kpi_card, bar_chart, comparison_bars, highlight_cards,
        )

        template_path = r"E:\PPT-Design-Skill\docs\分析脚本\template.pptx"
        if not os.path.isfile(template_path):
            pytest.skip("Real template not available")

        result = analyze(template_path)
        assert "Size:" in result
        assert "#2E6504" in result

        prs = Presentation(template_path)
        original_count = len(prs.slides)

        s = add_slide(prs)
        rect(s, 0, 0, 13.333, 7.5, 'background', C=C)
        top_bar(s, 'accent', C=C)
        page_header(s, '乡村振兴核心数据总览', '2024年度关键指标汇总', C=C)

        kpi_w = (12.0 - 0.35 * 3) / 4
        kpi_data = [
            ('12.8亿', '年度农业总产值', '同比 +8.3%', True),
            ('1,247万', '新增就业人数', '同比 +15.2%', True),
            ('3,580', '示范村建设数', '同比 +22.6%', True),
            ('96.5%', '基础设施覆盖率', '同比 +3.1%', True),
        ]
        kx = 0.65
        for num, label, trend, up in kpi_data:
            kpi_card(s, kx, 1.55, kpi_w, 1.35, num, label, trend, up, C=C)
            kx += kpi_w + 0.35

        bar_chart(s, 1.65, 3.55, [
            ('东部地区', 0.92, '460亿'),
            ('中部地区', 0.75, '375亿'),
            ('西部地区', 0.68, '340亿'),
            ('东北地区', 0.48, '240亿'),
        ], max_width=5.0, C=C)

        s2 = add_slide(prs)
        rect(s2, 0, 0, 13.333, 7.5, 'background', C=C)
        top_bar(s2, 'accent', C=C)
        page_header(s2, '乡村振兴成效对比分析', '2023 vs 2024', C=C)

        comparison_bars(s2, 1.65, 1.95, [
            ('农业产值', '11.8亿', '12.8亿', 0.85, 0.92),
            ('就业人数', '1083万', '1247万', 0.65, 0.75),
            ('示范村数', '2920', '3580', 0.55, 0.68),
        ], max_width=4.0, C=C)

        highlight_cards(s2, 0.65, 5.0, [
            ('乡村旅游增速领跑', '同比增长34.2%', 'primary'),
            ('电商覆盖突破', '村级电商服务站覆盖率达78.5%', 'accent'),
            ('绿色转型加速', '清洁能源应用比例提升至42%', 'muted'),
        ], total_width=12.0, C=C)

        output_path = str(tmp_path / "output.pptx")
        prs.save(output_path)

        assert os.path.isfile(output_path)
        out_prs = Presentation(output_path)
        assert len(out_prs.slides) == original_count + 2
        assert len(out_prs.slides[-1].shapes) > 5
        assert len(out_prs.slides[-2].shapes) > 5

    def test_copy_decorations_preserves_template_look(self, tmp_path, C):
        from ppt_pro_max.build_helpers import add_slide, rect, top_bar, page_header, copy_decorations

        template_path = r"E:\PPT-Design-Skill\docs\分析脚本\template.pptx"
        if not os.path.isfile(template_path):
            pytest.skip("Real template not available")

        prs = Presentation(template_path)
        original_count = len(prs.slides)

        content_slide = None
        for slide in prs.slides:
            if len(slide.shapes) >= 5:
                content_slide = slide
                break
        assert content_slide is not None

        s = add_slide(prs)
        rect(s, 0, 0, 13.333, 7.5, 'background', C=C)
        copy_decorations(s, content_slide, skip_long_text=True, skip_image=True)
        page_header(s, '新页面标题', C=C)

        output_path = str(tmp_path / "with_decorations.pptx")
        prs.save(output_path)

        assert os.path.isfile(output_path)
        out_prs = Presentation(output_path)
        assert len(out_prs.slides) == original_count + 1
        new_slide = out_prs.slides[-1]
        assert len(new_slide.shapes) >= 3

# 兼容性脚本

早期仓库曾经从 `scripts/` 提供渲染和检查命令。现在维护中的实现位于
`skill/scripts/`，因为这些脚本属于随技能安装的文件包。

使用方式：

```powershell
python skill/scripts/check_runtime.py
python skill/scripts/inspect_pptx.py path/to/deck.pptx --pretty
powershell -ExecutionPolicy Bypass -File skill/scripts/render_pptx.ps1 `
  -InFile path/to/deck.pptx -OutDir output/rendered
```

早期引擎专用的 `generate_ppt.py`、`audit_pptx.py` 和 `layout_report.py` 不再
复制。这些脚本依赖 `ppt_pro_max`；现在由 `pptx-designer` 负责生成，最终的
视觉验收由 LLM 直接检查 PNG 完成。

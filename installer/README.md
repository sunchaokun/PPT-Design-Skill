# PPT Design Skill 安装器

安装器会将完整的 `skill/` 文件包（包括参考资料和渲染脚本）复制到 AI
编码工具使用的标准 `ppt-design-skill/` 目录中。如果没有提供 `--no-pip`，
同时会安装已经发布的 `pptx-designer` Python 软件包。

## 开始安装

执行任何安装命令前，请先克隆仓库：

```powershell
git clone https://github.com/sunchaokun/PPT-Design-Skill.git
cd PPT-Design-Skill
```

## 自动安装

```powershell
python installer/install.py --all --force
```

## 主要平台

```powershell
python installer/install.py --platform claude --force
python installer/install.py --platform codex --force
python installer/install.py --platform opencode --force
python installer/install.py --platform deepseek-harness --force
```

如需安装到项目目录，请使用 `--project --target C:\path\to\project`。
不使用 `--project` 时，安装器会安装到检测到的全局技能目录。

DeepSeek Harness 使用其文档规定的 `skills/<name>/SKILL.md` 文件包结构，支持
全局 `~/.dsh/skills` 和项目级 `.dsh/skills` 目录。Codex 使用共享的
`~/.agents/skills` 全局目录和项目级 `.codex/skills` 目录。

## 渲染依赖

```powershell
python installer/install.py --render-deps
```

该命令会在 Windows 上明确调用 `winget` 安装 LibreOffice 和 Poppler。
安装器不会在后台静默安装桌面应用程序。

如需只检查环境而不修改任何内容：

```powershell
python installer/install.py --check
```

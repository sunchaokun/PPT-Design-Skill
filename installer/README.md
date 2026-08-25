# PPT Design Skill installer

The installer copies the complete `skill/` bundle, including references and
render scripts, into the standard `ppt-design-skill/` directory for an AI
coding assistant.
It also installs the published `pptx-designer` Python package unless
`--no-pip` is provided.

## Automatic installation

```powershell
python installer/install.py --all --force
```

## Main platforms

```powershell
python installer/install.py --platform claude --force
python installer/install.py --platform codex --force
python installer/install.py --platform opencode --force
python installer/install.py --platform deepseek-harness --force
```

Use `--project --target C:\path\to\project` for project-local installation.
Without `--project`, the installer targets detected global skill roots.

DeepSeek Harness uses its documented `skills/<name>/SKILL.md` bundle layout and
supports the global `~/.dsh/skills` and project `.dsh/skills` roots. Codex uses
the shared `~/.agents/skills` global root and `.codex/skills` project root.

## Render dependencies

```powershell
python installer/install.py --render-deps
```

This explicitly uses `winget` on Windows to install LibreOffice and Poppler.
The installer never silently installs desktop applications.

Check without changing anything:

```powershell
python installer/install.py --check
```

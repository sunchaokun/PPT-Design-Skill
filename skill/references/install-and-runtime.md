# Installation and runtime

## Python package

The skill depends on the published `pptx-designer` package:

```powershell
python -m pip install --upgrade pptx-designer
```

Core rendering dependencies are installed with the package. Optional image
search or generation features may require:

```powershell
python -m pip install --upgrade "pptx-designer[images]"
python -m pip install --upgrade "pptx-designer[ai-images]"
```

Use `skill/scripts/check_runtime.py` before a generation task. It reports
missing Python or rendering dependencies without silently changing the user's
environment.

## Image generation credentials

Image generation is optional. When a build needs AI-generated images, place a
`.env` file in the presentation project that runs `build.py`, not in the
installed skill directory or Python site-packages. Copy the repository's
`.env.example` and configure one provider:

```powershell
Copy-Item .env.example .env
```

The package loads the nearest project `.env` while walking upward from the
current working directory. Existing process environment variables take
priority. The most explicit form is:

```dotenv
PPT_IMAGE_LLM_PROVIDER=gpt-image
OPENAI_API_KEY=your-api-key
# OPENAI_IMAGE_MODEL=gpt-image-1
# OPENAI_BASE_URL=https://api.openai.com/v1
```

Supported provider-specific keys include `ARK_API_KEY` for Seedream,
`GEMINI_API_KEY` for Gemini, `DASHSCOPE_API_KEY` for Wanx, and
`MOONSHOT_API_KEY` for Kimi. A generic OpenAI-compatible endpoint can use
`PPT_IMAGE_LLM_API_KEY`, `PPT_IMAGE_LLM_BASE_URL`, and
`PPT_IMAGE_LLM_MODEL`.

If no image API is configured, the skill should use supplied local images,
available stock-image search, or a host-provided image-generation callback;
it must not invent or expose credentials. Test the configuration with:

```powershell
pptx-designer image "editorial fashion portrait in a white atelier" --image-mode auto -v
```

## PPTX → PDF → PNG dependencies

The preferred Windows renderer is Microsoft PowerPoint through COM. The
headless fallback requires:

- LibreOffice (`soffice` or `soffice.bin`)
- Poppler (`pdftoppm`)

These are operating-system dependencies, not Python dependencies. Install
them through the user's approved system package mechanism and then rerun the
runtime check. The repository installer offers an explicit Windows-only
`python installer/install.py --platform all --force --render-deps` option using `winget`; it is never enabled
implicitly.

## Fonts

Fonts are part of visual correctness. Detect the target language and use fonts
available on the target machine. If a requested font is unavailable, report
the substitution before final delivery and inspect the PNG again.

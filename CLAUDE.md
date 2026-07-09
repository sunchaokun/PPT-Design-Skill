# CLAUDE.md

## PPT Design Skill

AI-powered PPT generation with 40,000+ style combinations.

## Quick Reference

```bash
# Generate PPT
python -m ppt_pro_max "topic" --style "warm fintech"

# Test
python -m pytest tests/ -q
```

## Architecture

4-phase: StoryPlanner → DesignDecider → ContentGenerator → PPTRenderer

## Key Rules

- Use `_add_picture_cover()` for images (Pillow pre-crop, no distortion)
- Cache-first for all image API calls
- Never modify ui-ux-pro-max code
- Windows: use `python` not `python3`

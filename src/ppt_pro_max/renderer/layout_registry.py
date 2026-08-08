"""Layout Registry — 12 master layouts with professional inch coordinates."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5

MARGIN = {"top": 0.6, "bottom": 0.6, "left": 0.9, "right": 0.9}
CONTENT_WIDTH = SLIDE_WIDTH - MARGIN["left"] - MARGIN["right"]

_TITLE_STD = {"x": 0.9, "y": 0.5, "width": 11.533, "height": 0.7, "font_size": 28, "font_weight": "bold", "alignment": "left", "color_role": "foreground", "font_role": "heading"}
_TITLE_COMPACT = {"x": 0.9, "y": 0.3, "width": 11.533, "height": 0.5, "font_size": 22, "font_weight": "bold", "alignment": "left", "color_role": "foreground", "font_role": "heading"}
_DIVIDER_STD = {"x": 0.9, "y": 1.25, "width": 2.0, "height": 0.04, "type": "decoration", "decoration_type": "title_underline"}
_HALF_COL = CONTENT_WIDTH / 2 - 0.1

MASTER_LAYOUTS: dict[str, dict[str, Any]] = {
    "title-slide": {
        "id": 0,
        "name": "Title Slide",
        "goal_mapping": ["hook"],
        "placeholders": {
            "hero_image": {
                "x": 0, "y": 0, "width": 13.333, "height": 7.5,
                "type": "image",
            },
            "title": {
                "x": 1.0, "y": 2.2, "width": 11.333, "height": 1.4,
                "font_size": 44, "font_weight": "bold", "alignment": "center",
                "color_role": "on-primary", "font_role": "heading",
            },
            "subtitle": {
                "x": 2.0, "y": 3.8, "width": 9.333, "height": 0.7,
                "font_size": 18, "font_weight": "normal", "alignment": "center",
                "color_role": "on-primary", "font_role": "body",
            },
            "date": {
                "x": 4.5, "y": 6.3, "width": 4.333, "height": 0.4,
                "font_size": 12, "font_weight": "normal", "alignment": "center",
                "color_role": "on-primary", "font_role": "body",
            },
        },
    },
    "section-header": {
        "id": 1,
        "name": "Section Header",
        "goal_mapping": ["section"],
        "placeholders": {
            "section_number": {
                "x": 5.167, "y": 1.8, "width": 3.0, "height": 1.0,
                "font_size": 72, "font_weight": "bold", "alignment": "center",
                "color_role": "accent", "font_role": "heading",
            },
            "section_title": {
                "x": 1.5, "y": 3.2, "width": 10.333, "height": 1.0,
                "font_size": 36, "font_weight": "semibold", "alignment": "center",
                "color_role": "foreground", "font_role": "heading",
            },
        },
    },
    "content-with-title": {
        "id": 2,
        "name": "Content With Title",
        "goal_mapping": ["problem", "solution", "features", "content", "market", "business", "team", "financial", "agitation", "proof"],
        "placeholders": {
            "accent_bar": {
                "x": 0, "y": 0, "width": 0.08, "height": 7.5,
                "type": "decoration", "decoration_type": "left_accent",
            },
            "title": dict(_TITLE_STD),
            "divider": dict(_DIVIDER_STD),
            "body": {
                "x": 0.9, "y": 1.6, "width": 7.0, "height": 5.0,
                "font_size": 16, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
            "side_visual": {
                "x": 8.5, "y": 1.6, "width": 4.0, "height": 5.0,
                "type": "image",
            },
        },
    },
    "two-column": {
        "id": 3,
        "name": "Two Column",
        "goal_mapping": ["comparison", "before-after"],
        "placeholders": {
            "title": dict(_TITLE_STD),
            "divider": dict(_DIVIDER_STD),
            "left_col": {
                "x": 0.9, "y": 1.6, "width": 5.5, "height": 5.0,
                "font_size": 16, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
            "right_col": {
                "x": 6.933, "y": 1.6, "width": 5.5, "height": 5.0,
                "font_size": 16, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "three-column-cards": {
        "id": 4,
        "name": "Three Column Cards",
        "goal_mapping": ["features", "solution"],
        "placeholders": {
            "title": dict(_TITLE_STD),
            "divider": dict(_DIVIDER_STD),
            "card1": {
                "x": 0.9, "y": 1.7, "width": 3.644, "height": 4.8,
                "font_size": 14, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
            "card2": {
                "x": 4.844, "y": 1.7, "width": 3.644, "height": 4.8,
                "font_size": 14, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
            "card3": {
                "x": 8.789, "y": 1.7, "width": 3.644, "height": 4.8,
                "font_size": 14, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
        },
    },
    "four-metrics": {
        "id": 5,
        "name": "Four Metrics",
        "goal_mapping": ["traction", "metrics"],
        "placeholders": {
            "title": dict(_TITLE_STD),
            "divider": dict(_DIVIDER_STD),
            "metric1": {
                "x": 0.9, "y": 1.8, "width": 2.708, "height": 3.2,
                "font_size": 40, "font_weight": "bold", "alignment": "center",
                "color_role": "accent", "font_role": "heading",
            },
            "metric2": {
                "x": 4.108, "y": 1.8, "width": 2.708, "height": 3.2,
                "font_size": 40, "font_weight": "bold", "alignment": "center",
                "color_role": "accent", "font_role": "heading",
            },
            "metric3": {
                "x": 7.317, "y": 1.8, "width": 2.708, "height": 3.2,
                "font_size": 40, "font_weight": "bold", "alignment": "center",
                "color_role": "accent", "font_role": "heading",
            },
            "metric4": {
                "x": 10.525, "y": 1.8, "width": 2.708, "height": 3.2,
                "font_size": 40, "font_weight": "bold", "alignment": "center",
                "color_role": "accent", "font_role": "heading",
            },
        },
    },
    "big-number": {
        "id": 6,
        "name": "Big Number",
        "goal_mapping": ["agitation", "big-stat"],
        "placeholders": {
            "big_number": {
                "x": 2.0, "y": 1.5, "width": 9.333, "height": 2.5,
                "font_size": 96, "font_weight": "bold", "alignment": "center",
                "color_role": "accent", "font_role": "heading",
            },
            "label": {
                "x": 2.0, "y": 4.3, "width": 9.333, "height": 0.8,
                "font_size": 22, "font_weight": "normal", "alignment": "center",
                "color_role": "foreground", "font_role": "body",
            },
            "body": {
                "x": 2.5, "y": 5.3, "width": 8.333, "height": 1.5,
                "font_size": 14, "font_weight": "normal", "alignment": "center",
                "color_role": "muted-foreground", "font_role": "body",
            },
        },
    },
    "quote": {
        "id": 7,
        "name": "Quote",
        "goal_mapping": ["testimonial", "social-proof", "testimonials"],
        "placeholders": {
            "quote_mark": {
                "x": 1.0, "y": 1.5, "width": 1.5, "height": 1.5,
                "font_size": 80, "font_weight": "bold", "alignment": "left",
                "color_role": "accent", "font_role": "heading",
            },
            "quote_text": {
                "x": 1.5, "y": 2.3, "width": 10.333, "height": 2.5,
                "font_size": 22, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
            "quote_author": {
                "x": 1.5, "y": 5.2, "width": 10.333, "height": 0.6,
                "font_size": 14, "font_weight": "semibold", "alignment": "left",
                "color_role": "muted-foreground", "font_role": "body",
            },
        },
    },
    "chart-focus": {
        "id": 8,
        "name": "Chart Focus",
        "goal_mapping": ["data", "financial"],
        "placeholders": {
            "title": dict(_TITLE_STD),
            "divider": dict(_DIVIDER_STD),
            "chart": {
                "x": 1.2, "y": 1.6, "width": 7.5, "height": 5.0,
                "type": "chart",
            },
            "insight": {
                "x": 9.2, "y": 1.8, "width": 3.5, "height": 4.5,
                "font_size": 14, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "image-plus-text": {
        "id": 9,
        "name": "Image + Text",
        "goal_mapping": ["vision", "demo", "product"],
        "placeholders": {
            "image": {
                "x": 0, "y": 0, "width": 7.5, "height": 7.5,
                "type": "image",
            },
            "title": {
                "x": 8.0, "y": 1.5, "width": 4.8, "height": 1.0,
                "font_size": 28, "font_weight": "bold", "alignment": "left",
                "color_role": "foreground", "font_role": "heading",
            },
            "body": {
                "x": 8.0, "y": 2.8, "width": 4.8, "height": 3.5,
                "font_size": 16, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "cta-closing": {
        "id": 10,
        "name": "CTA Closing",
        "goal_mapping": ["cta", "offer", "pricing"],
        "placeholders": {
            "hero_image": {
                "x": 0, "y": 0, "width": 13.333, "height": 7.5,
                "type": "image",
            },
            "title": {
                "x": 1.5, "y": 2.0, "width": 10.333, "height": 1.2,
                "font_size": 44, "font_weight": "bold", "alignment": "center",
                "color_role": "on-primary", "font_role": "heading",
            },
            "subtitle": {
                "x": 2.5, "y": 3.5, "width": 8.333, "height": 0.7,
                "font_size": 18, "font_weight": "normal", "alignment": "center",
                "color_role": "on-primary", "font_role": "body",
            },
            "cta_button": {
                "x": 4.667, "y": 4.8, "width": 4.0, "height": 0.8,
                "font_size": 18, "font_weight": "bold", "alignment": "center",
                "color_role": "on-primary", "font_role": "body",
                "bg_color_role": "accent",
            },
        },
    },
    "blank": {
        "id": 11,
        "name": "Blank",
        "goal_mapping": ["custom"],
        "placeholders": {},
    },
    "grid-2x2-cards": {
        "id": 12,
        "name": "Grid 2x2 Cards",
        "goal_mapping": ["features", "overview", "education"],
        "placeholders": {
            "title": {
                "x": 0.9, "y": 0.3, "width": 11.533, "height": 0.6,
                "font_size": 24, "font_weight": "bold", "alignment": "left",
                "color_role": "foreground", "font_role": "heading",
            },
            "card1": {
                "x": 0.9, "y": 1.2, "width": 5.5, "height": 2.8,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
            "card2": {
                "x": 6.933, "y": 1.2, "width": 5.5, "height": 2.8,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
            "card3": {
                "x": 0.9, "y": 4.3, "width": 5.5, "height": 2.8,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
            "card4": {
                "x": 6.933, "y": 4.3, "width": 5.5, "height": 2.8,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
        },
    },
    "dense-bullets": {
        "id": 13,
        "name": "Dense Bullets",
        "goal_mapping": ["education", "training", "concept", "review"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "body": {
                "x": 0.9, "y": 1.0, "width": 11.533, "height": 6.0,
                "font_size": 11, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "two-column-dense": {
        "id": 14,
        "name": "Two Column Dense",
        "goal_mapping": ["comparison", "education", "analysis"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "left_col": {
                "x": 0.9, "y": 1.0, "width": 5.5, "height": 6.0,
                "font_size": 11, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
            "right_col": {
                "x": 6.933, "y": 1.0, "width": 5.5, "height": 6.0,
                "font_size": 11, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "table-layout": {
        "id": 15,
        "name": "Table Layout",
        "goal_mapping": ["report", "metrics", "analysis", "financial"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "body": {
                "x": 0.9, "y": 1.0, "width": 11.533, "height": 6.0,
                "font_size": 10, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "sidebar-left": {
        "id": 16,
        "name": "Sidebar Left",
        "goal_mapping": ["navigation", "agenda", "overview"],
        "placeholders": {
            "sidebar": {
                "x": 0, "y": 0, "width": 3.5, "height": 7.5,
                "font_size": 14, "font_weight": "normal", "alignment": "left",
                "color_role": "on-primary", "font_role": "body",
            },
            "title": {
                "x": 4.0, "y": 0.5, "width": 8.433, "height": 0.7,
                "font_size": 28, "font_weight": "bold", "alignment": "left",
                "color_role": "foreground", "font_role": "heading",
            },
            "body": {
                "x": 4.0, "y": 1.5, "width": 8.433, "height": 5.5,
                "font_size": 14, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "exercise-layout": {
        "id": 17,
        "name": "Exercise Layout",
        "goal_mapping": ["exercise", "practice", "training"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "instructions": {
                "x": 0.9, "y": 1.0, "width": 11.533, "height": 1.5,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "muted-foreground", "font_role": "body",
            },
            "body": {
                "x": 0.9, "y": 2.8, "width": 11.533, "height": 4.2,
                "font_size": 11, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "code-block": {
        "id": 18,
        "name": "Code Block",
        "goal_mapping": ["code", "technical", "demo"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "code": {
                "x": 0.9, "y": 1.0, "width": 11.533, "height": 5.5,
                "font_size": 10, "font_weight": "normal", "alignment": "left",
                "color_role": "on-primary", "font_role": "body",
            },
            "note": {
                "x": 0.9, "y": 6.7, "width": 11.533, "height": 0.5,
                "font_size": 10, "font_weight": "normal", "alignment": "left",
                "color_role": "muted-foreground", "font_role": "body",
            },
        },
    },
    "timeline-horizontal": {
        "id": 19,
        "name": "Timeline Horizontal",
        "goal_mapping": ["timeline", "process", "roadmap"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "body": {
                "x": 0.9, "y": 1.2, "width": 11.533, "height": 5.5,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "swot-matrix": {
        "id": 20,
        "name": "SWOT Matrix",
        "goal_mapping": ["swot", "strategy", "analysis"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "card1": {
                "x": 0.9, "y": 1.2, "width": 5.5, "height": 2.8,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
            "card2": {
                "x": 6.933, "y": 1.2, "width": 5.5, "height": 2.8,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
            "card3": {
                "x": 0.9, "y": 4.3, "width": 5.5, "height": 2.8,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
            "card4": {
                "x": 6.933, "y": 4.3, "width": 5.5, "height": 2.8,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
                "type": "card",
            },
        },
    },
    "funnel": {
        "id": 21,
        "name": "Funnel",
        "goal_mapping": ["funnel", "pipeline", "conversion"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "body": {
                "x": 0.9, "y": 1.2, "width": 11.533, "height": 5.5,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
    "cycle-diagram": {
        "id": 22,
        "name": "Cycle Diagram",
        "goal_mapping": ["cycle", "loop", "iteration"],
        "placeholders": {
            "title": dict(_TITLE_COMPACT),
            "body": {
                "x": 0.9, "y": 1.2, "width": 11.533, "height": 5.5,
                "font_size": 12, "font_weight": "normal", "alignment": "left",
                "color_role": "foreground", "font_role": "body",
            },
        },
    },
}


class LayoutRegistry:
    def __init__(self):
        self._layouts = dict(MASTER_LAYOUTS)
        self._load_csv_overrides()

    def get_layout(self, layout_name: str) -> dict[str, Any] | None:
        return self._layouts.get(layout_name)

    def get_layout_by_goal(self, goal: str) -> dict[str, Any] | None:
        for layout in self._layouts.values():
            if goal in layout.get("goal_mapping", []):
                return layout
        return self._layouts.get("content-with-title")

    def list_layouts(self) -> list[str]:
        return list(self._layouts.keys())

    def _load_csv_overrides(self) -> None:
        csv_path = self._find_data_file("ppt-master-layouts.csv")
        if csv_path is None:
            return

        try:
            with open(csv_path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    layout_name = row.get("layout_name", "")
                    ph_name = row.get("placeholder_name", "")
                    if not layout_name or not ph_name:
                        continue

                    if layout_name not in self._layouts:
                        self._layouts[layout_name] = {
                            "id": int(row.get("layout_id", 0)),
                            "name": layout_name,
                            "goal_mapping": row.get("goal_mapping", "").split(";"),
                            "placeholders": {},
                        }

                    self._layouts[layout_name]["placeholders"][ph_name] = {
                        "x": float(row.get("x", 0)),
                        "y": float(row.get("y", 0)),
                        "width": float(row.get("width", 0)),
                        "height": float(row.get("height", 0)),
                        "font_size": int(row.get("font_size", 16)),
                        "font_weight": row.get("font_weight", "normal"),
                        "alignment": row.get("alignment", "left"),
                        "color_role": row.get("color_role", "foreground"),
                        "font_role": row.get("font_role", "body"),
                    }
        except Exception:
            pass

    @staticmethod
    def _find_data_file(filename: str) -> Path | None:
        pkg_dir = Path(__file__).parent
        for candidate in [
            pkg_dir.parent.parent.parent / "data" / "ppt" / filename,
            pkg_dir.parent.parent / "data" / "ppt" / filename,
            pkg_dir.parent / "data" / "ppt" / filename,
        ]:
            if candidate.exists():
                return candidate

        try:
            import importlib.resources as pkg_resources
            ref = pkg_resources.files("ppt_pro_max") / ".." / "data" / "ppt" / filename
            with pkg_resources.as_file(ref) as path:
                if path.exists():
                    return Path(path)
        except Exception:
            pass

        return None

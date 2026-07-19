"""BrandSpec — brand specification data structure."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional


@dataclass
class BrandSpec:
    source: str = "none"

    colors: Optional[dict[str, str]] = None
    fonts: Optional[dict[str, str]] = None
    spacing: Optional[dict[str, Any]] = None
    logo: Optional[dict[str, Any]] = None
    layout_mapping: Optional[dict[str, int]] = None
    template_layouts: Optional[list[dict[str, Any]]] = None
    dark_mode: bool = False
    footer: Optional[dict[str, Any]] = None
    watermark: Optional[dict[str, Any]] = None
    _dna_actual_colors: Optional[dict[str, int]] = None

    @classmethod
    def from_brand_json(cls, data: dict[str, Any]) -> BrandSpec:
        return cls(
            source="brand_json",
            colors=data.get("colors"),
            fonts=data.get("fonts"),
            spacing=data.get("spacing"),
            logo=data.get("logo"),
            layout_mapping=data.get("layout_mapping"),
            dark_mode=bool(data.get("dark_mode", False)) or data.get("color_scheme") == "dark",
            footer=data.get("footer"),
            watermark=data.get("watermark"),
        )

    @classmethod
    def merge(cls, template_spec: BrandSpec, brand_data: dict[str, Any]) -> BrandSpec:
        brand_override = cls.from_brand_json(brand_data)
        merged_colors = dict(template_spec.colors) if template_spec.colors else {}
        if brand_override.colors:
            merged_colors.update(brand_override.colors)

        merged_fonts = dict(template_spec.fonts) if template_spec.fonts else {}
        if brand_override.fonts:
            merged_fonts.update(brand_override.fonts)

        return cls(
            source="merged",
            colors=merged_colors or None,
            fonts=merged_fonts or None,
            spacing=brand_override.spacing or template_spec.spacing,
            logo=brand_override.logo or template_spec.logo,
            layout_mapping=brand_override.layout_mapping or template_spec.layout_mapping,
            template_layouts=template_spec.template_layouts,
            dark_mode=brand_override.dark_mode or template_spec.dark_mode,
            footer=brand_override.footer or template_spec.footer,
            watermark=brand_override.watermark or template_spec.watermark,
        )

    @classmethod
    def merge_template_priority(cls, template_spec: BrandSpec, brand_data: dict[str, Any]) -> BrandSpec:
        brand_override = cls.from_brand_json(brand_data)

        merged_colors = dict(template_spec.colors) if template_spec.colors else {}
        if brand_override.colors:
            for k, v in brand_override.colors.items():
                if k not in merged_colors:
                    merged_colors[k] = v

        merged_fonts = dict(template_spec.fonts) if template_spec.fonts else {}
        if brand_override.fonts:
            for k, v in brand_override.fonts.items():
                if k not in merged_fonts:
                    merged_fonts[k] = v

        return cls(
            source="merged",
            colors=merged_colors or None,
            fonts=merged_fonts or None,
            spacing=brand_override.spacing or template_spec.spacing,
            logo=brand_override.logo or template_spec.logo,
            layout_mapping=brand_override.layout_mapping or template_spec.layout_mapping,
            template_layouts=template_spec.template_layouts,
            dark_mode=brand_override.dark_mode or template_spec.dark_mode,
            footer=brand_override.footer or template_spec.footer,
            watermark=brand_override.watermark or template_spec.watermark,
        )

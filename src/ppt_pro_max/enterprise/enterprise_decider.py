"""EnterpriseDesignDecider — brand-constrained design decisions."""

from __future__ import annotations

from ppt_pro_max.enterprise.brand_spec import BrandSpec


_BUSINESS_MODE_DEFAULT_DENSITY = {
    "pitch": 4,
    "education": 7,
    "training": 6,
    "report": 8,
}


class EnterpriseDesignDecider:

    def __init__(
        self,
        brand_spec: BrandSpec | None = None,
        business_mode: str | None = None,
    ):
        self.brand_spec = brand_spec or BrandSpec()
        self.business_mode = business_mode or "pitch"

    def resolve_layout_index(self, goal: str) -> int | None:
        if self.brand_spec.layout_mapping:
            return self.brand_spec.layout_mapping.get(goal)
        return None

    def suggest_density(self) -> int:
        return _BUSINESS_MODE_DEFAULT_DENSITY.get(self.business_mode, 4)

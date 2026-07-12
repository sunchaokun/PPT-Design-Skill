"""DataSplitter — split diagram data across pages when content overflows."""

from __future__ import annotations

from typing import Any


def split_data(
    data: dict[str, Any],
    max_items_per_page: int,
) -> list[dict[str, Any]]:
    nodes = data.get("nodes", [])
    stages = data.get("stages", [])
    events = data.get("events", [])

    if nodes:
        return _split_list(data, "nodes", nodes, max_items_per_page)
    if stages:
        return _split_list(data, "stages", stages, max_items_per_page)
    if events:
        return _split_list(data, "events", events, max_items_per_page)

    return [data]


def _split_list(
    original: dict[str, Any],
    key: str,
    items: list[dict[str, Any]],
    max_per_page: int,
) -> list[dict[str, Any]]:
    if len(items) <= max_per_page:
        return [original]

    result: list[dict[str, Any]] = []
    for i in range(0, len(items), max_per_page):
        chunk = items[i : i + max_per_page]
        page_data = {k: v for k, v in original.items() if k != key}
        page_data[key] = chunk
        page_data["page_index"] = i // max_per_page + 1
        page_data["total_pages"] = (len(items) + max_per_page - 1) // max_per_page
        result.append(page_data)
    return result

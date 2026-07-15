"""Batch import script — builds the component library from PPT素材库.

Features:
- Progress file for resume after interruption
- Gzipped XML storage for disk efficiency
- Batch transaction for speed
- Skip node_count=0 groups (pure decorative, no text to fill)
- Detailed log output

Usage:
    python -m ppt_pro_max.scripts.build_library [--materials-dir DIR] [--db-path PATH] [--batch-size N]

Can be run in background:
    start /B python -m ppt_pro_max.scripts.build_library > build_library.log 2>&1
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import time

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

_DEFAULT_MATERIALS = r"E:\BaiduNetdiskDownload\ABC"
_DEFAULT_DB = r"E:\PPT-Design-Skill\component_library\index.db"
_DEFAULT_BATCH = 10


def _collect_pptx(materials_dir: str) -> list[str]:
    pptx_files = []
    for root, dirs, files in os.walk(materials_dir):
        for f in sorted(files):
            if f.endswith(".pptx") and not f.startswith("~"):
                pptx_files.append(os.path.join(root, f))
    return pptx_files


def _load_progress(progress_path: str) -> set[int]:
    if not os.path.isfile(progress_path):
        return set()
    try:
        with open(progress_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return set(data.get("completed_indices", []))
    except Exception:
        return set()


def _save_progress(progress_path: str, completed: set[int]) -> None:
    with open(progress_path, "w", encoding="utf-8") as f:
        json.dump({"completed_indices": sorted(completed), "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")}, f)


def build_library(
    materials_dir: str = _DEFAULT_MATERIALS,
    db_path: str = _DEFAULT_DB,
    batch_size: int = _DEFAULT_BATCH,
    min_node_count: int = 1,
) -> dict:
    from ppt_pro_max.enterprise.component_library import ComponentLibrary

    progress_path = os.path.join(os.path.dirname(db_path), "import_progress.json")

    pptx_files = _collect_pptx(materials_dir)
    log.info("Found %d PPT files in %s", len(pptx_files), materials_dir)

    completed = _load_progress(progress_path)
    if completed:
        log.info("Resuming: %d files already processed", len(completed))

    lib = ComponentLibrary(db_path=db_path)

    total_added = 0
    total_skipped = 0
    total_errors = 0
    total_time = 0.0

    pending = [(i, p) for i, p in enumerate(pptx_files) if i not in completed]

    for batch_start in range(0, len(pending), batch_size):
        batch = pending[batch_start:batch_start + batch_size]
        batch_paths = [p for _, p in batch]
        batch_indices = [i for i, _ in batch]

        t0 = time.time()
        try:
            result = lib.bulk_import(batch_paths, min_node_count=min_node_count)
            elapsed = time.time() - t0
        except Exception as e:
            log.error("Batch failed: %s", e)
            total_errors += len(batch)
            total_time += time.time() - t0
            continue

        total_added += result["added"]
        total_skipped += result["skipped"]
        total_errors += result["errors"]
        total_time += elapsed

        completed.update(batch_indices)
        _save_progress(progress_path, completed)

        stats = lib.stats()
        log.info(
            "[%d/%d] added=%d skip=%d err=%d time=%.1fs total=%d",
            len(completed), len(pptx_files),
            result["added"], result["skipped"], result["errors"],
            elapsed, stats["total"],
        )

    stats = lib.stats()
    log.info("=== BUILD COMPLETE ===")
    log.info("Files processed: %d/%d", len(completed), len(pptx_files))
    log.info("Components added: %d", total_added)
    log.info("Skipped (dedup/empty): %d", total_skipped)
    log.info("Errors: %d", total_errors)
    log.info("Total time: %.1fs (%.1fs/file)", total_time, total_time / max(len(completed), 1))
    log.info("Library: %s", stats)

    for cat in ["infographic", "process", "hierarchy", "swot", "chart", "timeline", "comparison", "features", "dashboard"]:
        cov = lib.coverage(cat)
        if cov:
            total_in_cat = sum(cov.values())
            log.info("  %s: %d components", cat, total_in_cat)

    lib.close()

    if len(completed) == len(pptx_files):
        os.remove(progress_path)
        log.info("Progress file cleaned up")

    return {
        "files_processed": len(completed),
        "total_files": len(pptx_files),
        "added": total_added,
        "skipped": total_skipped,
        "errors": total_errors,
        "time": total_time,
        "stats": stats,
    }


def main():
    parser = argparse.ArgumentParser(description="Build component library from PPT素材库")
    parser.add_argument("--materials-dir", default=_DEFAULT_MATERIALS, help="PPT素材库 directory")
    parser.add_argument("--db-path", default=_DEFAULT_DB, help="SQLite database path")
    parser.add_argument("--batch-size", type=int, default=_DEFAULT_BATCH, help="Files per batch transaction")
    parser.add_argument("--min-node-count", type=int, default=1, help="Minimum text nodes to import (0=import all)")
    args = parser.parse_args()

    build_library(
        materials_dir=args.materials_dir,
        db_path=args.db_path,
        batch_size=args.batch_size,
        min_node_count=args.min_node_count,
    )


if __name__ == "__main__":
    main()

"""Validate visual-capability metadata without requiring jsonschema."""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
SCHEMA_DIR = ROOT / "skill" / "schemas"
INDEXES = {
    "packs": ROOT / "skill" / "references" / "visual-rendering-packs" / "rendering_packs_index.json",
    "compositions": ROOT / "skill" / "references" / "composition-recipes" / "composition_index.json",
    "prototypes": ROOT / "skill" / "templates" / "case-prototype-index.json",
}


class ValidationError(ValueError):
    pass


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ValidationError(f"missing file: {path}") from exc
    except json.JSONDecodeError as exc:
        raise ValidationError(f"invalid JSON in {path}: {exc}") from exc


def check_schema(value: Any, schema: dict[str, Any], path: str = "$", root: dict[str, Any] | None = None) -> None:
    root = root or schema
    if "$ref" in schema:
        ref = schema["$ref"]
        if not ref.startswith("#/"):
            raise ValidationError(f"{path}: unsupported ref {ref}")
        target: Any = root
        for part in ref[2:].split("/"):
            target = target[part]
        check_schema(value, target, path, root)
        return
    if "enum" in schema and value not in schema["enum"]:
        raise ValidationError(f"{path}: expected one of {schema['enum']}, got {value!r}")
    typ = schema.get("type")
    type_ok = {
        "object": isinstance(value, dict),
        "array": isinstance(value, list),
        "string": isinstance(value, str),
        "integer": isinstance(value, int) and not isinstance(value, bool),
        "number": isinstance(value, (int, float)) and not isinstance(value, bool),
        "boolean": isinstance(value, bool),
    }
    if typ and not type_ok.get(typ, True):
        raise ValidationError(f"{path}: expected {typ}")
    if isinstance(value, str):
        if len(value) < schema.get("minLength", 0):
            raise ValidationError(f"{path}: string is too short")
        if "pattern" in schema and not re.search(schema["pattern"], value):
            raise ValidationError(f"{path}: does not match {schema['pattern']}")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value < schema.get("minimum", value) or value > schema.get("maximum", value):
            raise ValidationError(f"{path}: number outside range")
        if "exclusiveMinimum" in schema and value <= schema["exclusiveMinimum"]:
            raise ValidationError(f"{path}: number is not above exclusive minimum")
    if isinstance(value, list):
        if len(value) < schema.get("minItems", 0):
            raise ValidationError(f"{path}: too few items")
        if "items" in schema:
            for i, item in enumerate(value):
                check_schema(item, schema["items"], f"{path}[{i}]", root)
    if isinstance(value, dict):
        for key in schema.get("required", []):
            if key not in value:
                raise ValidationError(f"{path}: missing required field {key!r}")
        if schema.get("additionalProperties") is False:
            allowed = set(schema.get("properties", {}))
            extra = sorted(set(value) - allowed)
            if extra:
                raise ValidationError(f"{path}: unexpected fields {extra}")
        for key, item in value.items():
            if key in schema.get("properties", {}):
                check_schema(item, schema["properties"][key], f"{path}.{key}", root)


def relative_inside(case_root: Path, rel: str) -> Path:
    if Path(rel).is_absolute() or ".." in Path(rel).parts:
        raise ValidationError(f"path must be relative and cannot contain ..: {rel}")
    resolved = (case_root / rel).resolve()
    if case_root.resolve() not in resolved.parents and resolved != case_root.resolve():
        raise ValidationError(f"path escapes case root: {rel}")
    return resolved


def pptx_slide_count(path: Path) -> int:
    """Count presentation slides without importing the optional PPTX stack."""
    try:
        with zipfile.ZipFile(path) as archive:
            return sum(
                1
                for name in archive.namelist()
                if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
            )
    except (FileNotFoundError, zipfile.BadZipFile) as exc:
        raise ValidationError(f"invalid PPTX source: {path}") from exc


def validate_prototype_record(record: dict[str, Any], record_label: str, record_path: Path) -> None:
    """Validate the evidence chain behind a prototype record."""
    case_root = ROOT / record["case_root"]
    if not case_root.is_dir():
        raise ValidationError(f"{record_label}: missing case root {record['case_root']}")
    for field in ("source_paths", "preview_paths", "recipe_paths", "object_map_paths"):
        for rel in record[field]:
            resolved = relative_inside(case_root, rel)
            if not resolved.is_file():
                raise ValidationError(f"{record_label}.{field}: missing file {rel}")
    source_pptx = [relative_inside(case_root, rel) for rel in record["source_paths"] if rel.lower().endswith(".pptx")]
    if not source_pptx:
        raise ValidationError(f"{record_label}.source_paths: at least one PPTX is required")
    slide_count = pptx_slide_count(source_pptx[0])
    for slide_id in record["slide_ids"]:
        match = re.fullmatch(r"P(\d+)", slide_id)
        if not match or int(match.group(1)) > slide_count:
            raise ValidationError(f"{record_label}.slide_ids: {slide_id} is outside PPTX slide count {slide_count}")
    if record.get("license_status") != "verified" or record.get("context_allowed") is not True:
        if record_label.startswith("prototype["):
            raise ValidationError(f"{record_label}: unverified or disallowed evidence must remain blocked")
    if not record_path.is_file():
        raise ValidationError(f"{record_label}: missing record file {record_path}")


def validate_indexes() -> list[str]:
    errors: list[str] = []
    loaded: dict[str, dict[str, Any]] = {}
    for label, path in INDEXES.items():
        try:
            index = load_json(path)
            if not isinstance(index, dict) or index.get("schema_version") != 1:
                raise ValidationError("expected object with schema_version 1")
            expected = "packs" if label == "packs" else "entries" if label == "compositions" else "prototypes"
            if not isinstance(index.get(expected), list):
                raise ValidationError(f"{expected} must be an array")
            ids = []
            for item in index[expected]:
                if isinstance(item, dict):
                    ids.append(item.get("pack_id", item.get("composition_id", item.get("prototype_id"))))
            if len(ids) != len(set(ids)):
                raise ValidationError(f"{expected} contains duplicate IDs")
            if label in {"packs", "compositions"}:
                schema_name = "rendering-pack.schema.json" if label == "packs" else "composition-entry.schema.json"
                schema = load_json(SCHEMA_DIR / schema_name)
                for i, item in enumerate(index[expected]):
                    if not isinstance(item, dict) or not isinstance(item.get("path"), str):
                        raise ValidationError(f"{expected}[{i}] must contain a path")
                    asset_path = (ROOT / item["path"]).resolve()
                    if ROOT.resolve() not in asset_path.parents:
                        raise ValidationError(f"{expected}[{i}] path escapes repository: {item['path']}")
                    asset = load_json(asset_path)
                    check_schema(asset, schema, f"{expected}[{i}].asset", schema)
                    if asset.get("pack_id", asset.get("composition_id")) != item.get("pack_id", item.get("composition_id")):
                        raise ValidationError(f"{expected}[{i}] index id does not match asset")
                    loaded.setdefault(label, {})[asset.get("pack_id", asset.get("composition_id"))] = asset
            else:
                loaded[label] = {str(item.get("prototype_id")): item for item in index[expected] if isinstance(item, dict) and item.get("prototype_id")}
                blocked_ids = {item.get("prototype_id") for item in index.get("blocked", []) if isinstance(item, dict)}
                if blocked_ids & set(loaded[label]):
                    raise ValidationError("blocked prototype ID is also registered as active")
                prototype_schema = load_json(SCHEMA_DIR / "prototype-record.schema.json")
                for i, item in enumerate(index.get(expected, [])):
                    if not isinstance(item, dict) or not item.get("path"):
                        continue
                    record_path = (ROOT / item["path"]).resolve()
                    if ROOT.resolve() not in record_path.parents:
                        raise ValidationError(f"prototypes[{i}] path escapes repository")
                    record = load_json(record_path)
                    check_schema(record, prototype_schema, f"prototypes[{i}].record", prototype_schema)
                    validate_prototype_record(record, f"prototype[{i}]", record_path)
                for i, item in enumerate(index.get("blocked", [])):
                    if not isinstance(item, dict) or not isinstance(item.get("path"), str):
                        raise ValidationError(f"blocked[{i}] must contain a metadata path")
                    record_path = (ROOT / item["path"]).resolve()
                    if ROOT.resolve() not in record_path.parents:
                        raise ValidationError(f"blocked[{i}] path escapes repository")
                    record = load_json(record_path)
                    check_schema(record, prototype_schema, f"blocked[{i}].record", prototype_schema)
                    if record.get("prototype_id") != item.get("prototype_id") or record.get("case_id") != item.get("case_id"):
                        raise ValidationError(f"blocked[{i}] summary does not match record")
                    validate_prototype_record(record, f"blocked[{i}]", record_path)
        except ValidationError as exc:
            errors.append(f"{path}: {exc}")
    if not errors:
        pack_index = load_json(INDEXES["packs"])
        composition_index = load_json(INDEXES["compositions"])
        composition_ids = {item.get("composition_id") for item in composition_index.get("entries", []) if isinstance(item, dict)}
        prototype_ids = {item.get("prototype_id") for item in load_json(INDEXES["prototypes"]).get("prototypes", []) if isinstance(item, dict)}
        for pack in loaded.get("packs", {}).values():
            missing = set(pack.get("composition_families", [])) - composition_ids
            if missing:
                errors.append(f"pack {pack.get('pack_id')}: missing composition references {sorted(missing)}")
        for entry in loaded.get("compositions", {}).values():
            missing = set(entry.get("prototype_ids", [])) - prototype_ids
            if missing:
                errors.append(f"composition {entry.get('composition_id')}: missing prototype references {sorted(missing)}")
    return errors


def validate_schema_files() -> list[str]:
    errors: list[str] = []
    for path in sorted(SCHEMA_DIR.glob("*.schema.json")):
        try:
            schema = load_json(path)
            if schema.get("$schema") != "https://json-schema.org/draft/2020-12/schema":
                raise ValidationError("not Draft 2020-12")
            if schema.get("type") != "object":
                raise ValidationError("root must be an object schema")
        except ValidationError as exc:
            errors.append(f"{path}: {exc}")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args()
    errors = validate_schema_files() + validate_indexes()
    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1
    print("OK: schema files and empty/runtime indexes are valid")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

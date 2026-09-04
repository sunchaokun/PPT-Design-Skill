from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def run_script(name: str, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(ROOT / "skill/scripts" / name), *args],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
    )


def test_visual_schemas_and_empty_indexes_are_valid() -> None:
    result = run_script("validate_visual_pack.py")
    assert result.returncode == 0, result.stdout + result.stderr


def test_asset_inventory_is_machine_readable() -> None:
    result = run_script("inspect_visual_assets.py")
    assert result.returncode == 0, result.stdout + result.stderr
    assert "packs=" in result.stdout
    assert "registered_cases=6" in result.stdout
    assert "status=inventory_only" in result.stdout


def test_acceptance_record_enforces_score_and_status_semantics(tmp_path: Path) -> None:
    criteria = [
        {"id": "visual_thesis", "name": "Visual thesis", "score": 2, "evidence": ["p01"]},
        {"id": "composition", "name": "Composition", "score": 2, "evidence": ["p01"]},
        {"id": "material", "name": "Material", "score": 2, "evidence": ["p01"]},
        {"id": "readability", "name": "Readability", "score": 2, "evidence": ["p01"]},
        {"id": "editability", "name": "Editability", "score": 2, "evidence": ["inspect"]},
        {"id": "reproducibility", "name": "Reproducibility", "score": 1, "evidence": ["build"]},
    ]
    record = tmp_path / "acceptance.json"
    value = {"record_id": "demo", "target_id": "demo-p01", "criteria": criteria, "score": 11,
             "reviewer": "reviewer-a", "evidence": ["p01"], "status": "PASS"}
    record.write_text(json.dumps(value), encoding="utf-8")
    result = run_script("validate_acceptance_record.py", str(record))
    assert result.returncode == 0, result.stdout + result.stderr
    value["score"] = 10
    record.write_text(json.dumps(value), encoding="utf-8")
    result = run_script("validate_acceptance_record.py", str(record))
    assert result.returncode == 1
    assert "criteria sum" in result.stdout


def test_runtime_trace_rejects_inconsistent_gate_states(tmp_path: Path) -> None:
    trace = tmp_path / "trace.json"
    value = {
        "run_id": "run-1", "mode": "build", "prototype_ids": [], "recipe_ids": [],
        "seed": 17, "package_version": "1.0.0b10", "p01_gate": "PASS",
        "final_visual_gate": "PASS",
    }
    trace.write_text(json.dumps(value), encoding="utf-8")
    result = run_script("validate_runtime_trace.py", str(trace))
    assert result.returncode == 0, result.stdout + result.stderr
    value["p01_gate"] = "BLOCKED"
    trace.write_text(json.dumps(value), encoding="utf-8")
    result = run_script("validate_runtime_trace.py", str(trace))
    assert result.returncode == 1
    assert "cannot PASS" in result.stdout


def test_regression_audit_does_not_invent_missing_baseline(tmp_path: Path) -> None:
    result = run_script("audit_regression_pairs.py")
    assert result.returncode == 1
    report = json.loads(result.stdout)
    assert report["status"] == "BLOCKED"
    assert len(report["entries"]) == 3
    assert all("missing baseline PPTX" in item["blocking_reasons"] for item in report["entries"])


def test_regression_audit_rejects_self_comparison(tmp_path: Path) -> None:
    import zipfile

    pptx = tmp_path / "deck.pptx"
    with zipfile.ZipFile(pptx, "w") as archive:
        archive.writestr("ppt/slides/slide1.xml", "<slide/>")
    manifest = tmp_path / "manifest.json"
    entries = []
    for domain in ("technical", "scientific", "brand_architecture"):
        entries.append({"regression_id": domain, "domain": domain, "case_id": domain,
                        "baseline": "deck.pptx", "upgraded": "deck.pptx",
                        "status": "PASS", "reason": ""})
    manifest.write_text(json.dumps({"schema_version": 1, "entries": entries}), encoding="utf-8")
    result = run_script("audit_regression_pairs.py", "--root", str(tmp_path), "--manifest", str(manifest))
    assert result.returncode == 1
    report = json.loads(result.stdout)
    assert all(any("different PPTX files" in reason for reason in item["blocking_reasons"]) for item in report["entries"])


def test_installer_check_uses_runtime_dependency_resolution() -> None:
    result = subprocess.run(
        [sys.executable, str(ROOT / "installer/install.py"), "--check"],
        cwd=ROOT, text=True, capture_output=True, check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "OK libreoffice" in result.stdout
    assert "OK poppler" in result.stdout


def test_provenance_audit_reports_hashes_without_overclaiming() -> None:
    result = run_script("audit_prototype_provenance.py")
    assert result.returncode == 0, result.stdout + result.stderr
    report = json.loads(result.stdout)
    assert report["status"] == "UNVERIFIED"
    assert len(report["records"]) == 8
    assert all(item["status"] == "UNVERIFIED" for item in report["records"])
    assert all(item["hashes"].get("build") for item in report["records"])


def test_case_output_audit_reports_render_readiness_without_overclaiming() -> None:
    result = run_script("audit_case_outputs.py")
    assert result.returncode == 0
    report = json.loads(result.stdout)
    assert len(report["cases"]) == 6
    statuses = {item["status"] for item in report["cases"]}
    assert statuses == {"render_ready"}


def test_workflow_state_machine_supports_resume(tmp_path: Path) -> None:
    state = tmp_path / "state.json"
    for action in ("route", "directions", "confirm", "p01", "pass", "review", "pass"):
        result = run_script("run_visual_workflow.py", str(state), action)
        assert result.returncode == 0, result.stdout + result.stderr
    data = json.loads(state.read_text(encoding="utf-8"))
    assert data["status"] == "PASS"
    assert len(data["history"]) == 7


def test_workflow_rejects_invalid_transition(tmp_path: Path) -> None:
    state = tmp_path / "state.json"
    result = run_script("run_visual_workflow.py", str(state), "pass")
    assert result.returncode == 1
    assert "invalid" in result.stdout


def test_refresh_writes_atomic_state_for_registered_case(tmp_path: Path) -> None:
    case = tmp_path / "cases" / "demo-case"
    case.mkdir(parents=True)
    (case / "marker.txt").write_text("v1", encoding="utf-8")
    index = tmp_path / "index.json"
    index.write_text(json.dumps({"schema_version": 1, "cases_root": "cases", "prototypes": [{
        "prototype_id": "demo-case-p01", "case_id": "demo-case", "case_root": "cases/demo-case"
    }]}), encoding="utf-8")
    cache = tmp_path / "cache.json"
    lock = tmp_path / "cache.lock"
    result = run_script("refresh_case_prototypes.py", "--root", str(tmp_path), "--index", str(index), "--cache", str(cache), "--lock", str(lock))
    assert result.returncode == 0, result.stdout + result.stderr
    state = json.loads(cache.read_text(encoding="utf-8"))
    assert state["records"]["demo-case-p01"]["status"] == "valid"
    assert state["records"]["demo-case-p01"]["change"] == "new"
    result = run_script("refresh_case_prototypes.py", "--root", str(tmp_path), "--index", str(index), "--cache", str(cache), "--lock", str(lock))
    assert result.returncode == 0, result.stdout + result.stderr
    state = json.loads(cache.read_text(encoding="utf-8"))
    assert state["records"]["demo-case-p01"]["change"] == "unchanged"
    (case / "marker.txt").write_text("v2", encoding="utf-8")
    result = run_script("refresh_case_prototypes.py", "--root", str(tmp_path), "--index", str(index), "--cache", str(cache), "--lock", str(lock))
    assert result.returncode == 0, result.stdout + result.stderr
    state = json.loads(cache.read_text(encoding="utf-8"))
    assert state["records"]["demo-case-p01"]["change"] == "changed"
    assert not lock.exists()


def test_refresh_keeps_blocked_prototypes_out_of_valid_runtime_state(tmp_path: Path) -> None:
    case = tmp_path / "cases" / "blocked-case"
    case.mkdir(parents=True)
    index = tmp_path / "index.json"
    index.write_text(json.dumps({
        "schema_version": 1,
        "cases_root": "cases",
        "prototypes": [],
        "blocked": [{
            "prototype_id": "blocked-case-p01",
            "case_id": "blocked-case",
            "case_root": "cases/blocked-case",
            "reason": "permission pending",
        }],
    }), encoding="utf-8")
    cache = tmp_path / "cache.json"
    result = run_script(
        "refresh_case_prototypes.py",
        "--root", str(tmp_path), "--index", str(index),
        "--cache", str(cache), "--lock", str(tmp_path / "cache.lock"),
    )
    assert result.returncode == 0, result.stdout + result.stderr
    state = json.loads(cache.read_text(encoding="utf-8"))
    assert state["records"]["blocked-case-p01"]["status"] == "blocked"
    assert state["records"]["blocked-case-p01"]["blocked"] is True


def test_refresh_fails_cleanly_when_lock_is_busy(tmp_path: Path) -> None:
    index = tmp_path / "index.json"
    index.write_text(json.dumps({"schema_version": 1, "cases_root": "cases", "prototypes": []}), encoding="utf-8")
    lock = tmp_path / "cache.lock"
    lock.write_text("busy", encoding="utf-8")
    result = run_script("refresh_case_prototypes.py", "--root", str(tmp_path), "--index", str(index), "--cache", str(tmp_path / "cache.json"), "--lock", str(lock))
    assert result.returncode == 1
    assert "busy" in result.stderr or "busy" in result.stdout


def test_vi_build_delivery_removes_template_pages_and_runs_qa(tmp_path: Path) -> None:
    from pptx_designer import Presentation
    from pptx_designer.enterprise import VIBuildDelivery

    template = ROOT / "examples/new_examplex/louvre_abudhabi/output/louvre_abudhabi_complete.pptx"
    presentation = Presentation(template_path=str(template))

    class FixtureAdapter:
        def render(self, spec, prs):
            return prs.slides.add_slide(prs.slide_layouts[6])

    delivery = VIBuildDelivery(presentation, FixtureAdapter())
    delivery.add({"delivery_origin": "build_components", "page_role": "content"})
    output = tmp_path / "vi-delivery.pptx"
    report = delivery.finalize(str(output), check_overlaps=True)
    assert output.exists()
    assert report.status == "pass"


def test_vi_template_adapter_rebinds_confirmed_framework_slot(tmp_path: Path) -> None:
    from pptx_designer import Presentation
    from pptx_designer.enterprise import VIBuildDelivery, VITemplateAdapter

    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.shapes.add_textbox(100000, 100000, 1000000, 300000).text = "PLACEHOLDER"
    context = {
        "framework_pages": [{"id": "title-1", "role": "title", "reference_slide": 1,
                             "text_contract": {"strict": True, "clear_shape_indices": []}}],
        "content_slots": [{"id": "title", "page_role": "title", "target": {"shape_index": 0}}],
        "locks": [], "source": {"template_fingerprint": "fixture"}, "visual_grammar": {},
    }
    adapter = VITemplateAdapter(context)
    spec = adapter.compile(page_role="title", content={"slots": {"title": "Confirmed title"}})
    delivery = VIBuildDelivery(presentation, adapter)
    delivery.add(spec)
    report = delivery.finalize(str(tmp_path / "framework-rebind.pptx"), check_overlaps=False)
    assert report.status == "pass"


def test_vi_template_adapter_rejects_unowned_content_page() -> None:
    from pptx_designer.enterprise import VITemplateAdapter

    adapter = VITemplateAdapter({"framework_pages": [], "content_slots": [], "locks": []})
    try:
        adapter.compile(page_role="content")
    except ValueError as exc:
        assert "compile_atomic" in str(exc)
    else:
        raise AssertionError("content page must require an atomic Build plan")

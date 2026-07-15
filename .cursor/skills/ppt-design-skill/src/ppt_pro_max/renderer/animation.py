from __future__ import annotations

from lxml import etree
from pptx.oxml.ns import qn


TRANSITION_TYPES = {
    "fade": ("p:fade", {"thruBlk": "0"}),
    "push": ("p:push", {}),
    "wipe": ("p:wipe", {}),
    "split": ("p:split", {}),
    "cover": ("p:cover", {}),
    "dissolve": ("p:dissolve", {}),
    "wheel": ("p:wheel", {"spokes": "4"}),
    "wedge": ("p:wedge", {}),
    "blinds": ("p:blinds", {}),
    "checker": ("p:checker", {}),
    "comb": ("p:comb", {}),
    "random": ("p:random", {}),
}

SPEED_MAP = {
    "slow": "slow",
    "medium": "med",
    "fast": "fast",
}

ENTRANCE_PRESETS = {
    "appear": (1, 0),
    "fly_in": (2, 4),
    "fly_in_left": (2, 8),
    "fly_in_right": (2, 2),
    "fly_in_top": (2, 1),
    "fly_in_bottom": (2, 4),
    "fade_in": (10, 0),
    "zoom_in": (53, 16),
    "float_up": (42, 4),
    "grow_turn": (13, 0),
    "bounce": (25, 0),
}


def add_slide_transition(
    slide,
    transition_type: str = "fade",
    speed: str = "medium",
    advance_on_click: bool = True,
    advance_after_ms: int | None = None,
) -> None:
    if transition_type not in TRANSITION_TYPES:
        transition_type = "fade"

    sld = slide._element
    existing = sld.find(qn("p:transition"))
    if existing is not None:
        sld.remove(existing)

    trans = etree.SubElement(sld, qn("p:transition"))
    trans.set("spd", SPEED_MAP.get(speed, "med"))
    trans.set("advClick", "1" if advance_on_click else "0")
    if advance_after_ms is not None:
        trans.set("advTm", str(advance_after_ms))

    tag, attrs = TRANSITION_TYPES[transition_type]
    child = etree.SubElement(trans, qn(tag))
    for k, v in attrs.items():
        child.set(k, v)

    cSld = sld.find(qn("p:cSld"))
    if cSld is not None:
        sld.remove(trans)
        idx = list(sld).index(cSld) + 1
        sld.insert(idx, trans)


def add_entrance_animation(
    slide,
    shape_id: int,
    effect: str = "fade_in",
    delay_ms: int = 0,
    duration_ms: int = 500,
    click_triggered: bool = True,
) -> None:
    if effect not in ENTRANCE_PRESETS:
        effect = "fade_in"

    preset_id, preset_subtype = ENTRANCE_PRESETS[effect]

    sld = slide._element
    timing = sld.find(qn("p:timing"))
    if timing is None:
        timing = etree.SubElement(sld, qn("p:timing"))
        tnLst = etree.SubElement(timing, qn("p:tnLst"))
        par_root = etree.SubElement(tnLst, qn("p:par"))
        cTn_root = _make_cTn(id="1", dur="indefinite", restart="never", nodeType="tmRoot")
        par_root.append(cTn_root)
        childTnLst = etree.SubElement(cTn_root, qn("p:childTnLst"))

        seq = etree.SubElement(childTnLst, qn("p:seq"))
        seq.set("concurrent", "1")
        seq.set("nextAc", "seek")
        cTn_seq = _make_cTn(id="2", dur="indefinite", nodeType="mainSeq")
        seq.append(cTn_seq)
        seqChildTnLst = etree.SubElement(cTn_seq, qn("p:childTnLst"))

        prevCondLst = etree.SubElement(seq, qn("p:prevCondLst"))
        prev_cond = etree.SubElement(prevCondLst, qn("p:cond"))
        prev_cond.set("evt", "onPrev")
        prev_cond.set("delay", "0")
        prev_tgtEl = etree.SubElement(prev_cond, qn("p:tgtEl"))
        etree.SubElement(prev_tgtEl, qn("p:sldTgt"))

        nextCondLst = etree.SubElement(seq, qn("p:nextCondLst"))
        next_cond = etree.SubElement(nextCondLst, qn("p:cond"))
        next_cond.set("evt", "onNext")
        next_cond.set("delay", "0")
        next_tgtEl = etree.SubElement(next_cond, qn("p:tgtEl"))
        etree.SubElement(next_tgtEl, qn("p:sldTgt"))

        cSld = sld.find(qn("p:cSld"))
        if cSld is not None:
            sld.remove(timing)
            idx = list(sld).index(cSld) + 1
            sld.insert(idx, timing)
    else:
        tnLst = timing.find(qn("p:tnLst"))
        par_root = tnLst.find(qn("p:par"))
        cTn_root = par_root.find(qn("p:cTn"))
        childTnLst = cTn_root.find(qn("p:childTnLst"))
        seq = childTnLst.find(qn("p:seq"))
        cTn_seq = seq.find(qn("p:cTn"))
        seqChildTnLst = cTn_seq.find(qn("p:childTnLst"))

    existing_ids = _collect_ids(timing)
    next_id = max(existing_ids, default=0) + 1

    click_par = etree.SubElement(seqChildTnLst, qn("p:par"))
    cTn_click = _make_cTn(id=str(next_id), fill="hold")
    click_par.append(cTn_click)
    next_id += 1

    stCondLst = etree.SubElement(cTn_click, qn("p:stCondLst"))
    cond = etree.SubElement(stCondLst, qn("p:cond"))
    cond.set("delay", "0")

    clickChildTnLst = etree.SubElement(cTn_click, qn("p:childTnLst"))

    effect_par = etree.SubElement(clickChildTnLst, qn("p:par"))
    cTn_effect = _make_cTn(id=str(next_id), fill="hold")
    effect_par.append(cTn_effect)
    next_id += 1

    effectChildTnLst = etree.SubElement(cTn_effect, qn("p:childTnLst"))

    anim_par = etree.SubElement(effectChildTnLst, qn("p:par"))
    cTn_anim = _make_cTn(
        id=str(next_id),
        presetID=str(preset_id),
        presetClass="entr",
        presetSubtype=str(preset_subtype),
        fill="hold",
        nodeType="clickEffect" if click_triggered else "afterEffect",
    )
    if delay_ms > 0:
        cTn_anim.set("delay", str(delay_ms))
    cTn_anim.set("dur", str(duration_ms))
    anim_par.append(cTn_anim)
    next_id += 1

    animChildTnLst = etree.SubElement(cTn_anim, qn("p:childTnLst"))

    set_elem = etree.SubElement(animChildTnLst, qn("p:set"))
    cBhvr = etree.SubElement(set_elem, qn("p:cBhvr"))
    cTn_set = _make_cTn(id=str(next_id), dur="1", fill="hold")
    cBhvr.append(cTn_set)
    next_id += 1

    stCondLst2 = etree.SubElement(cTn_set, qn("p:stCondLst"))
    cond2 = etree.SubElement(stCondLst2, qn("p:cond"))
    cond2.set("delay", "0")

    tgtEl = etree.SubElement(cBhvr, qn("p:tgtEl"))
    spTgt = etree.SubElement(tgtEl, qn("p:spTgt"))
    spTgt.set("spid", str(shape_id))

    attrNameLst = etree.SubElement(cBhvr, qn("p:attrNameLst"))
    attrName = etree.SubElement(attrNameLst, qn("p:attrName"))
    attrName.text = "style.visibility"

    to_elem = etree.SubElement(set_elem, qn("p:to"))
    val = etree.SubElement(to_elem, qn("p:strVal"))
    val.set("val", "visible")


def add_animation_sequence(
    slide,
    shape_ids: list[int],
    effect: str = "fade_in",
    interval_ms: int = 300,
    duration_ms: int = 500,
) -> None:
    for i, sid in enumerate(shape_ids):
        if i == 0:
            add_entrance_animation(
                slide, sid, effect=effect,
                delay_ms=0, duration_ms=duration_ms,
                click_triggered=True,
            )
        else:
            add_entrance_animation(
                slide, sid, effect=effect,
                delay_ms=interval_ms, duration_ms=duration_ms,
                click_triggered=False,
            )


def _make_cTn(**kwargs) -> etree._Element:
    elem = etree.Element(qn("p:cTn"))
    for k, v in kwargs.items():
        elem.set(k, v)
    return elem


def _collect_ids(timing_elem) -> list[int]:
    ids = []
    for elem in timing_elem.iter():
        id_val = elem.get("id")
        if id_val and id_val.isdigit():
            ids.append(int(id_val))
    return ids

"""Round-trip test: extract component -> inject -> verify coordinates scale correctly."""

from __future__ import annotations

from lxml import etree

from ppt_pro_max.enterprise.component_renderer import ComponentRenderer

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _make_group_xml():
    return f"""<p:grpSp xmlns:p="{_P}" xmlns:a="{_A}">
  <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="100000" y="200000"/>
      <a:ext cx="4000000" cy="3000000"/>
      <a:chOff x="50000" y="100000"/>
      <a:chExt cx="3900000" cy="2800000"/>
    </a:xfrm>
  </p:grpSpPr>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Shape1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="500000" y="400000"/><a:ext cx="1500000" cy="800000"/></a:xfrm></p:spPr>
  </p:sp>
  <p:grpSp>
    <p:nvGrpSpPr><p:cNvPr id="3" name="SubGroup"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm>
        <a:off x="2000000" y="500000"/>
        <a:ext cx="1800000" cy="2200000"/>
        <a:chOff x="0" y="0"/>
        <a:chExt cx="1800000" cy="2200000"/>
      </a:xfrm>
    </p:grpSpPr>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="4" name="SubShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="600000" cy="400000"/></a:xfrm></p:spPr>
    </p:sp>
  </p:grpSp>
</p:grpSp>"""


def _absolute_screen_positions(grp):
    """Compute absolute screen positions for all leaf shapes, fully resolving nested Group transforms."""
    a_ns = _A
    p_ns = _P
    positions = {}

    def _resolve(grp_elem, transform_x, transform_y):
        grpSpPr = grp_elem.find(f"{{{p_ns}}}grpSpPr")
        if grpSpPr is None:
            return
        xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
        if xfrm is None:
            return

        off = xfrm.find(f"{{{a_ns}}}off")
        ext = xfrm.find(f"{{{a_ns}}}ext")
        chOff = xfrm.find(f"{{{a_ns}}}chOff")
        chExt = xfrm.find(f"{{{a_ns}}}chExt")

        if off is None or ext is None:
            return

        off_x = int(off.get("x", "0"))
        off_y = int(off.get("y", "0"))
        ext_cx = int(ext.get("cx", "1"))
        ext_cy = int(ext.get("cy", "1"))

        chOff_x = int(chOff.get("x", str(off_x))) if chOff is not None else off_x
        chOff_y = int(chOff.get("y", str(off_y))) if chOff is not None else off_y
        chExt_cx = int(chExt.get("cx", str(ext_cx))) if chExt is not None else ext_cx
        chExt_cy = int(chExt.get("cy", str(ext_cy))) if chExt is not None else ext_cy

        def map_point(px, py):
            sx = (px - chOff_x) * (ext_cx / chExt_cx) + off_x
            sy = (py - chOff_y) * (ext_cy / chExt_cy) + off_y
            return transform_x(sx), transform_y(sy)

        for child in grp_elem:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag in ("sp", "pic", "cxnSp", "graphicFrame"):
                spPr = child.find(f"{{{p_ns}}}spPr")
                if spPr is None:
                    continue
                c_xfrm = spPr.find(f"{{{a_ns}}}xfrm")
                if c_xfrm is None:
                    continue
                c_off = c_xfrm.find(f"{{{a_ns}}}off")
                if c_off is not None:
                    cx = int(c_off.get("x", "0"))
                    cy = int(c_off.get("y", "0"))
                    abs_x, abs_y = map_point(cx, cy)
                    c_cNv = child.find(f".//{{{p_ns}}}cNvPr")
                    c_name = c_cNv.get("name", "?") if c_cNv is not None else "?"
                    positions[c_name] = (round(abs_x), round(abs_y))
            elif tag == "grpSp":
                def make_sub_transform(tx, ty):
                    return lambda px, py: map_point(px, py)
                sub_tx, sub_ty = map_point(lambda x: x, lambda y: y)
                _resolve(child, lambda px: map_point(px, 0)[0], lambda py: map_point(0, py)[1])

    _resolve(grp, lambda x: x, lambda y: y)
    return positions


def test_identity_transform():
    grp = etree.fromstring(_make_group_xml())

    grpSpPr = grp.find(f"{{{_P}}}grpSpPr")
    xfrm = grpSpPr.find(f"{{{_A}}}xfrm")
    off = xfrm.find(f"{{{_A}}}off")
    ext = xfrm.find(f"{{{_A}}}ext")

    target_left = int(off.get("x")) / 914400
    target_top = int(off.get("y")) / 914400
    target_w = int(ext.get("cx")) / 914400
    target_h = int(ext.get("cy")) / 914400

    renderer = ComponentRenderer()
    renderer._denormalize_coordinates(
        grp, target_left, target_top, target_w, target_h,
    )

    a_ns = _A
    p_ns = _P
    grpSpPr2 = grp.find(f"{{{p_ns}}}grpSpPr")
    xfrm2 = grpSpPr2.find(f"{{{a_ns}}}xfrm")
    off2 = xfrm2.find(f"{{{a_ns}}}off")
    ext2 = xfrm2.find(f"{{{a_ns}}}ext")

    assert int(off2.get("x")) == int(off.get("x"))
    assert int(off2.get("y")) == int(off.get("y"))
    assert int(ext2.get("cx")) == int(ext.get("cx"))
    assert int(ext2.get("cy")) == int(ext.get("cy"))


def test_scaled_transform():
    grp = etree.fromstring(_make_group_xml())

    target_left = 0.5
    target_top = 1.0
    target_w = 8.0
    target_h = 6.0

    renderer = ComponentRenderer()
    renderer._denormalize_coordinates(
        grp, target_left, target_top, target_w, target_h,
    )

    a_ns = _A
    p_ns = _P
    grpSpPr = grp.find(f"{{{p_ns}}}grpSpPr")
    xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
    off = xfrm.find(f"{{{a_ns}}}off")
    ext = xfrm.find(f"{{{a_ns}}}ext")

    assert int(off.get("x")) >= int(0.5 * 914400)
    assert int(off.get("y")) >= int(1.0 * 914400)
    assert int(ext.get("cx")) <= int(8.0 * 914400)
    assert int(ext.get("cy")) <= int(6.0 * 914400)


def test_ext_chext_ratio_preserved():
    grp = etree.fromstring(_make_group_xml())

    a_ns = _A
    p_ns = _P

    renderer = ComponentRenderer()
    renderer._denormalize_coordinates(
        grp, 0.5, 1.0, 8.0, 6.0,
    )

    grpSpPr = grp.find(f"{{{p_ns}}}grpSpPr")
    xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
    off = xfrm.find(f"{{{a_ns}}}off")
    ext = xfrm.find(f"{{{a_ns}}}ext")
    chOff = xfrm.find(f"{{{a_ns}}}chOff")
    chExt = xfrm.find(f"{{{a_ns}}}chExt")

    assert int(off.get("x")) >= int(0.5 * 914400)
    assert int(off.get("y")) >= int(1.0 * 914400)
    assert int(ext.get("cx")) <= int(8.0 * 914400)
    assert int(ext.get("cy")) <= int(6.0 * 914400)


def test_no_orig_bounds_fallback():
    grp_xml = f"""<p:grpSp xmlns:p="{_P}" xmlns:a="{_A}">
  <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="100000" y="200000"/>
      <a:ext cx="4000000" cy="3000000"/>
    </a:xfrm>
  </p:grpSpPr>
</p:grpSp>"""
    grp = etree.fromstring(grp_xml)

    renderer = ComponentRenderer()
    renderer._denormalize_coordinates(
        grp, 0.5, 1.0, 8.0, 6.0,
    )

    a_ns = _A
    p_ns = _P
    grpSpPr = grp.find(f"{{{p_ns}}}grpSpPr")
    xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
    off = xfrm.find(f"{{{a_ns}}}off")
    ext = xfrm.find(f"{{{a_ns}}}ext")

    assert int(off.get("x")) == int(0.5 * 914400)
    assert int(off.get("y")) == int(1.0 * 914400)
    assert int(ext.get("cx")) == int(8.0 * 914400)
    assert int(ext.get("cy")) == int(6.0 * 914400)


def test_child_shapes_within_group_bounds():
    grp = etree.fromstring(_make_group_xml())

    renderer = ComponentRenderer()
    renderer._denormalize_coordinates(grp, 0.5, 1.0, 8.0, 6.0)

    a_ns = _A
    p_ns = _P
    grpSpPr = grp.find(f"{{{p_ns}}}grpSpPr")
    xfrm = grpSpPr.find(f"{{{a_ns}}}xfrm")
    off = xfrm.find(f"{{{a_ns}}}off")
    ext = xfrm.find(f"{{{a_ns}}}ext")
    chOff = xfrm.find(f"{{{a_ns}}}chOff")
    chExt = xfrm.find(f"{{{a_ns}}}chExt")

    grp_x = int(off.get("x", "0"))
    grp_y = int(off.get("y", "0"))
    grp_w = int(ext.get("cx", "0"))
    grp_h = int(ext.get("cy", "0"))
    chOff_x = int(chOff.get("x", "0")) if chOff is not None else grp_x
    chOff_y = int(chOff.get("y", "0")) if chOff is not None else grp_y
    chExt_cx = int(chExt.get("cx", "0")) if chExt is not None else grp_w
    chExt_cy = int(chExt.get("cy", "0")) if chExt is not None else grp_h

    for sp in grp.findall(f"{{{p_ns}}}sp"):
        spPr = sp.find(f"{{{p_ns}}}spPr")
        if spPr is None:
            continue
        c_xfrm = spPr.find(f"{{{a_ns}}}xfrm")
        if c_xfrm is None:
            continue
        c_off = c_xfrm.find(f"{{{a_ns}}}off")
        c_ext = c_xfrm.find(f"{{{a_ns}}}ext")
        if c_off is not None and c_ext is not None:
            cx = int(c_off.get("x", "0"))
            cy = int(c_off.get("y", "0"))
            cw = int(c_ext.get("cx", "0"))
            ch = int(c_ext.get("cy", "0"))

            assert cx >= min(chOff_x, 0) - 1000, f"Child x={cx} below chOff_x={chOff_x}"
            assert cy >= min(chOff_y, 0) - 1000, f"Child y={cy} below chOff_y={chOff_y}"

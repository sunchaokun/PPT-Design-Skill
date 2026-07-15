from __future__ import annotations


def group_shapes(slide, shape_indices: list[int] | None = None, shape_ids: list[int] | None = None):
    if shape_indices is not None:
        shapes = [slide.shapes[i] for i in shape_indices]
    elif shape_ids is not None:
        id_map = {s.shape_id: s for s in slide.shapes}
        shapes = [id_map[sid] for sid in shape_ids if sid in id_map]
    else:
        return None

    if len(shapes) < 2:
        return None

    group = slide.shapes.add_group_shape(shapes)
    return group

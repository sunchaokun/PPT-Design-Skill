"""Tests for svg_compiler._affine — Affine transform + parse_transform."""

from ppt_pro_max.renderer.svg_compiler._affine import Affine, parse_transform


class TestAffineIdentity:
    def test_identity_apply(self):
        af = Affine()
        assert af.apply(3, 4) == (3.0, 4.0)

    def test_identity_compose(self):
        a = Affine()
        b = Affine()
        c = a.compose(b)
        assert c.a == 1.0 and c.d == 1.0 and c.e == 0.0 and c.f == 0.0


class TestAffineTranslate:
    def test_translate(self):
        af = Affine(1, 0, 0, 1, 10, 20)
        assert af.apply(0, 0) == (10.0, 20.0)

    def test_translate_compose(self):
        a = Affine(1, 0, 0, 1, 5, 0)
        b = Affine(1, 0, 0, 1, 0, 10)
        c = a.compose(b)
        assert c.apply(0, 0) == (5.0, 10.0)


class TestAffineScale:
    def test_scale_uniform(self):
        af = Affine(2, 0, 0, 2, 0, 0)
        assert af.apply(3, 4) == (6.0, 8.0)

    def test_scale_nonuniform(self):
        af = Affine(2, 0, 0, 3, 0, 0)
        assert af.apply(1, 1) == (2.0, 3.0)


class TestAffineRotate:
    def test_rotate_90(self):
        af = Affine(0, 1, -1, 0, 0, 0)
        x, y = af.apply(1, 0)
        assert abs(x - 0.0) < 1e-9
        assert abs(y - 1.0) < 1e-9

    def test_rotate_180(self):
        af = Affine(-1, 0, 0, -1, 0, 0)
        x, y = af.apply(1, 0)
        assert abs(x - (-1.0)) < 1e-9
        assert abs(y - 0.0) < 1e-9


class TestAffineCompose:
    def test_translate_then_scale(self):
        t = Affine(1, 0, 0, 1, 5, 0)
        s = Affine(2, 0, 0, 2, 0, 0)
        c = t.compose(s)
        assert c.apply(0, 0) == (10.0, 0.0)

    def test_scale_then_translate(self):
        s = Affine(2, 0, 0, 2, 0, 0)
        t = Affine(1, 0, 0, 1, 5, 0)
        c = s.compose(t)
        assert c.apply(0, 0) == (5.0, 0.0)


class TestParseTransform:
    def test_empty(self):
        af = parse_transform("")
        assert af.a == 1.0 and af.d == 1.0

    def test_none(self):
        af = parse_transform(None)
        assert af.a == 1.0 and af.d == 1.0

    def test_translate(self):
        af = parse_transform("translate(10,20)")
        assert af.apply(0, 0) == (10.0, 20.0)

    def test_translate_single_arg(self):
        af = parse_transform("translate(10)")
        assert af.apply(0, 0) == (10.0, 0.0)

    def test_scale(self):
        af = parse_transform("scale(2)")
        assert af.apply(3, 4) == (6.0, 8.0)

    def test_scale_xy(self):
        af = parse_transform("scale(2,3)")
        assert af.apply(1, 1) == (2.0, 3.0)

    def test_rotate_90(self):
        af = parse_transform("rotate(90)")
        x, y = af.apply(1, 0)
        assert abs(x) < 1e-6
        assert abs(y - 1.0) < 1e-6

    def test_rotate_with_center(self):
        af = parse_transform("rotate(90, 10, 10)")
        x, y = af.apply(10, 0)
        assert abs(x - (-20.0)) < 1e-6
        assert abs(y - 10.0) < 1e-6

    def test_matrix(self):
        af = parse_transform("matrix(2,0,0,3,10,20)")
        assert af.apply(0, 0) == (10.0, 20.0)
        assert af.apply(1, 1) == (12.0, 23.0)

    def test_nested(self):
        af = parse_transform("translate(10,0) scale(2)")
        assert af.apply(3, 4) == (26.0, 8.0)

    def test_skewX(self):
        af = parse_transform("skewX(45)")
        x, y = af.apply(0, 1)
        assert abs(x - 1.0) < 1e-6
        assert abs(y - 1.0) < 1e-6

    def test_skewY(self):
        af = parse_transform("skewY(45)")
        x, y = af.apply(1, 0)
        assert abs(x - 1.0) < 1e-6
        assert abs(y - 1.0) < 1e-6

    def test_unknown_op_ignored(self):
        af = parse_transform("foo(1,2)")
        assert af.a == 1.0 and af.d == 1.0

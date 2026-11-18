import unittest

from latex_cleaner import extract_main_formula


class LatexCleanerTests(unittest.TestCase):
    def test_preserves_multiline_matrix_and_font_commands(self):
        formula = r"\begin{array}{cc}\mathbb{R} & x^2 \\ \mathbf{v} & \frac{a}{b}\end{array}"
        self.assertEqual(extract_main_formula(formula), formula)

    def test_removes_only_paired_outer_delimiters(self):
        self.assertEqual(extract_main_formula(r"\[\frac{a}{b}\]"), r"\frac{a}{b}")
        self.assertEqual(extract_main_formula(r"\(x_i\)"), r"x_i")

    def test_keeps_unpaired_delimiter(self):
        self.assertEqual(extract_main_formula(r"$x_i"), r"$x_i")


if __name__ == '__main__':
    unittest.main()

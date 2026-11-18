"""Conservative cleanup of OCR LaTeX before copying it into Word.

An OCR result can contain matrices, multiple rows and font commands. Removing
those tokens changes the formula, so only remove a paired outer math wrapper.
"""


def clean_latex_for_word(latex_str):
    """Return the complete formula, without optional outer math delimiters."""
    formula = latex_str.strip()
    for opening, closing in ((r"\[", r"\]"), (r"\(", r"\)"), ("$$", "$$"), ("$", "$")):
        if formula.startswith(opening) and formula.endswith(closing):
            formula = formula[len(opening):-len(closing)].strip()
            break
    return formula


def extract_main_formula(latex_str):
    """Legacy name kept for callers; preserve every row of the formula."""
    return clean_latex_for_word(latex_str)

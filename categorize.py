"""
Map an invoice line-item description to the fixed category vocabulary
used in the 'Export Sales' sheet (e.g. "Full Shoes (Gents)").

The workbook's categories only depend on {Item Type, Gender}, not on the
material (Leather / Textile / Leather+Textile) and not on Grade 1 vs 2 --
Grade only adds the " Grade II" suffix.
"""
import re

TYPE_PATTERNS = [
    (r'full\s*shoes', 'Full Shoes'),
    (r'half\s*boots', 'Half Boots'),
    (r'sandals', 'Sandals'),
    (r'chappals?', 'Chappals'),
    (r'shoe\s*uppers?', 'Shoe Uppers'),
    (r'shoe\s*material', 'Shoe Material'),
    (r'finished\s*leather', 'Finished Leather'),
    (r'good\s*soles?', 'Good Soles'),
]

GENDER_PATTERNS = [
    (r'gents?', 'Gents'),
    (r'ladies', 'Ladies'),
]

# categories that never take a gender suffix
NO_GENDER_TYPES = {"Shoe Uppers", "Shoe Material", "Finished Leather", "Good Soles"}


def map_category(description: str, grade: str = "I") -> str:
    desc = description.lower()

    item_type = next((label for pat, label in TYPE_PATTERNS if re.search(pat, desc)), None)
    gender = next((label for pat, label in GENDER_PATTERNS if re.search(pat, desc)), None)

    if item_type is None:
        # Unknown item -> flag clearly instead of silently mis-filing it
        return f"UNMAPPED: {description}"

    if item_type in NO_GENDER_TYPES:
        category = item_type
    else:
        category = f"{item_type} ({gender or 'Gents'})"

    if grade == "II":
        category += " Grade II"

    return category

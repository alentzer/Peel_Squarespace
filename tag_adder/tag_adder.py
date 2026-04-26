import openpyxl
from openpyxl.styles import PatternFill

# -----------------------------
# FILE NAMES
# -----------------------------
MAPPING_FILE = "/workspaces/Peel_Squarespace/tag_adder/validation_list.xlsx"
PRODUCT_FILE = "/workspaces/Peel_Squarespace/tag_adder/space_items.xlsx"
OUTPUT_FILE = "space_items_UPDATED.xlsx"

# -----------------------------
# LOAD WORKBOOKS
# -----------------------------
mapping_wb = openpyxl.load_workbook(MAPPING_FILE)

# Use "validation_list" if it exists, otherwise default to the first sheet
if "validation_list" in mapping_wb.sheetnames:
    mapping_ws = mapping_wb["validation_list"]
else:
    mapping_ws = mapping_wb[mapping_wb.sheetnames[0]]
    print(f"⚠️ Sheet 'validation_list' not found. Using '{mapping_wb.sheetnames[0]}' instead.")

product_wb = openpyxl.load_workbook(PRODUCT_FILE)
product_ws = product_wb.active  # use first sheet; change if needed

# -----------------------------
# BUILD CATEGORY → TAGS LOOKUP
# -----------------------------
category_to_tags = {}

for row in mapping_ws.iter_rows(min_row=2, values_only=True):
    category, tags = row
    if category:
        cat_str = str(category).strip()
        tags_str = str(tags).strip() if tags else ""
        category_to_tags[cat_str] = tags_str

# -----------------------------
# YELLOW HIGHLIGHT STYLE
# -----------------------------
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# -----------------------------
# COLUMN INDEXES (1-based)
# -----------------------------
# Y = 25, Z = 26
CATEGORIES_COL = 25  # Y
TAGS_COL = 26        # Z

# -----------------------------
# PROCESS PRODUCTS
# -----------------------------
for row in range(2, product_ws.max_row + 1):
    categories_cell = product_ws.cell(row=row, column=CATEGORIES_COL)
    tags_cell = product_ws.cell(row=row, column=TAGS_COL)

    raw_categories = categories_cell.value

    # Skip rows where categories are missing or empty
    if not raw_categories or str(raw_categories).strip() == "":
        continue

    # Parse categories: "/art-books, /comics" -> ["/art-books", "/comics"]
    categories_list = [
        c.strip()
        for c in str(raw_categories).split(",")
        if c.strip() != ""
    ]

    merged_tags_list = []

    # Collect tags from all matched categories (strict match)
    for cat in categories_list:
        if cat in category_to_tags:
            tags_str = category_to_tags[cat]
            if tags_str:
                # Split mapped tags by comma and deduplicate
                for t in tags_str.split(","):
                    tag_clean = t.strip()
                    if tag_clean and tag_clean not in merged_tags_list:
                        merged_tags_list.append(tag_clean)

    # If no categories matched the mapping, do nothing
    if not merged_tags_list:
        continue

    # Build final tag string (comma-separated)
    final_tags = ", ".join(merged_tags_list)

    # Replace tags entirely based on category mapping
    tags_cell.value = final_tags
    tags_cell.fill = yellow_fill  # highlight changed cell

# -----------------------------
# SAVE UPDATED FILE
# -----------------------------
product_wb.save(OUTPUT_FILE)
print(f"✅ Done. Updated file saved as: {OUTPUT_FILE}")

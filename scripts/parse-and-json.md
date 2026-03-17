# parse-and-json.py — How It Works

## Test Template

The script is tested against `digital-templates/VSME-Digital-Template-Sample-1.1.1-unlocked.xlsx`, which is a manually unlocked version of the official VSME Digital Template Sample (version 1.1.1).

The following `template_*` defined names were **manually added** to the unlocked file to enable section detection for two sections that were missing named ranges in the original locked template:

| Added Defined Name | Sheet | Purpose |
|---|---|---|
| `template_c1_strategy_business_model_and_sustainability_if_applicable` | Social Disclosures | Section header for C1 — Strategy, Business Model and Sustainability |
| `template_b10_workforce_remuneration_collective_bargaining_and_training_always_reported` | Social Disclosures | Section header for B10 — Workforce Remuneration, Collective Bargaining and Training |

Without these additions, those two sections would not be detected by the section-discovery logic (which relies on `template_*` defined names pointing to bold header cells).

> **Note:** When updating to a future template version, check whether these section names have been officially added to the template. If so, the manually added names should be removed from the unlocked file to avoid duplicates.

---

## Overview

`parse-and-json.py` reads a VSME Digital Excel Template and outputs a structured JSON file intended as the data source for an online survey tool. It extracts:

- **XBRL fact values** already entered by the preparer
- **Survey field schemas** (labels, data types, input types, validation rules, enum options) for every expected field — including fields with no value yet
- **Yellow cell fields** — condition triggers and helper inputs (detected via fill colour `FFFFFF99`) that are not XBRL-reportable but essential for online survey conditional logic
- **Section structure** matching the visual layout of the Excel template
- **Multilingual translations** (9 EU languages)
- **Validation rules** derived from Excel IF-formula cells
- **Enum/dropdown option lists**
- **Template metadata and applicability flags**

---

## How the Excel Is Read

### Two Workbook Loads

The script opens the Excel file **twice**:

| Load | Mode | Purpose |
|------|------|---------|
| **First** (via `ExcelProcessor`) | `data_only=True` | Reads cached cell values (what the user entered). Used to extract XBRL facts and all `template_*`/`enum_*` named range values. |
| **Second** (inside `_extract_template_data()`) | `data_only=False` | Reads raw cell formulas. Used to parse validation IF-formulas, warning references, and yellow-cell label formulas. |

The second load is ephemeral — it is opened, scanned, and closed inside `_extract_template_data()`.

### XBRL Facts (First Load)

The base class `ExcelProcessor.populateReport()` handles XBRL fact extraction. It:

1. Resolves all non-`template_*`, non-`enum_*` defined names
2. Reads the cell values at those named ranges
3. Constructs typed `Fact` objects and attaches them to an `InlineReport`

The resulting `report._facts` list is what the script uses to populate field values in the JSON output.

### Named Ranges (First Load)

`getNamedRanges()` iterates all workbook defined names and reads their cell values. The script categorises each name:

| Prefix | Category | Used For |
|--------|----------|---------|
| `template_label_*` | `template_label` | Human-readable display labels |
| `template_*` (other) | `template` | Section headers, metadata, applicability flags |
| `enum_*` | `enum` | Dropdown option lists |
| everything else | `xbrl` | XBRL data fields |

### Translations Sheet (First Load)

The Translations sheet is read directly by cell coordinates:

- **Row 2, columns 3–K**: language codes (e.g. `en`, `da`, `fr`, `de`, `it`, `lt`, `pl`, `pt`, `es`)
- **Rows 4+**: `col A` = label ID, `cols C–K` = translated strings per language

All `template_label_*` keys are extracted into the `translations` dict.

### Section Structure (First Load)

Sections are discovered from `template_*` defined names whose target cell is **bold-formatted** in one of the four data sheets:

- `General Information`
- `Environmental Disclosures`
- `Social Disclosures`
- `Governance Disclosures`

Non-section `template_*` names are excluded via a hard-coded `_NON_SECTION_NAMES` set.

Applicability is determined by parsing the **evaluated cell text** (e.g. `[Always to be reported]`, `[If applicable]`) present in the resolved header value.

XBRL fields are assigned to sections by checking which XBRL defined-name cells fall within the section's row range (and column range for side-by-side sections). Fields are sorted by `(row, col)` to match the visual top-to-bottom order.

### Validation Rules (Second Load — Formula Scan)

The script scans every cell in all four data sheets for formulas referencing `template_label_ok`. These are the ~997 validation indicator cells. For each such cell it:

1. Identifies the **primary XBRL field** on the same row (searching left, then right, then by formula cell reference, then by `template_label_*` in column C as a fallback)
2. Detects **condition fields** referenced in the formula (named ranges, cross-sheet refs, local cell refs)
3. Classifies the rule as `required`, `conditional`, or `informational`

### Warning Definitions (Second Load — Formula Scan)

Cells matching `=template_label_*warning*` are extracted similarly. Each warning is associated with the XBRL field on the same row and attached to its validation rule entry.

### Yellow Cell Fields (Second Load — Fill Colour Scan)

Yellow cells (fill colour `FFFFFF99`) contain condition/trigger questions and helper inputs that are **not** XBRL-reportable but are essential for driving conditional logic in an online survey. They have no defined names, so the script discovers them by cell fill colour.

**Detection process:**

1. Iterates every cell in all four data sheets (second workbook load, `data_only=False`)
2. Checks `cell.fill.start_color.index == "FFFFFF99"`
3. For each yellow cell, classifies it as either:
   - **Label cell** — formula matches `re.compile(r"^=?template_label_")` → skipped (label-only)
   - **Value cell** — user-editable input → kept

**Label resolution:** For each value cell, the script searches the same row for a yellow label cell whose formula contains a `template_label_*` key. That key is used to resolve the English label and i18n translations.

**Filtering rules (value cells):**

| Condition | Action |
|-----------|--------|
| Formula containing cell references (e.g. `=IF(OR(G30...`) | Skipped — computed, not user-editable |
| Value is `"-"` or `"-"` | Skipped — placeholder dash |
| Literal boolean `True` / `False` | Kept — `inputType: "boolean"` |
| Data validation referencing `enum_*` formula | Kept — `inputType: "select"` with `options` |
| Numeric value | Kept — `inputType: "number"` |
| Other | Kept — `inputType: "text"` |

**DataType inference:**

| inputType | dataType |
|-----------|----------|
| `boolean` | `booleanItemType` |
| `number` | `decimalItemType` |
| `select` | `enumerationItemType` |
| `text` | `None` |

**Integration:** Yellow cell fields are interleaved with XBRL fields in the section's `fields` array, sorted by `(row, col)` to match the Excel visual order. They use `fieldId` (pattern: `yellow_{label_key}_{colLetter}{row}`) instead of `qname`, and carry `source: "yellowCell"`, `isRequired: false`, `isReportable: false`.

**Current counts (sample file):** 90 yellow cells scanned → 25 value fields extracted across 17 sections. The remaining 65 are label cells, computed formulas, or dash placeholders.

---

## Output Structure

```
report-facts.json
├── reportTitle
├── entityName
├── factCount
├── templateMetadata        ← resolved values of template_* names (non-label)
├── templateLabels          ← resolved values of template_label_* names
├── translations            ← i18n dict: label_key → {lang: text}
├── sectionApplicability    ← section_key → "always"|"conditional"|"optional"|…
├── omittedDisclosures      ← list of user-marked omitted sections
├── sections[]              ← one entry per Excel section header
│   ├── sectionId
│   ├── templateName
│   ├── sheet
│   ├── headerRow
│   ├── applicability
│   ├── hasFacts
│   └── fields[]            ← merged schema + data per field
│       ├── qname, label, labels (i18n)          ← XBRL fields
│       ├── fieldId, label, labels (i18n)        ← yellow cell fields (source: "yellowCell")
│       ├── dataType, periodType, inputType
│       ├── isRequired
│       ├── isReportable                         ← false for yellow cells
│       ├── validation       ← from IF-formula scan
│       ├── options / domainMembers
│       └── value            ← null if not filled in
├── enumLists               ← enum_* name → [option strings]
├── xbrlToEnum              ← XBRL field name → enum_* list name(s)
├── validationRules         ← top-level dict, same data as field.validation
├── warningDefinitions      ← warning_key → {lang: text}
└── definedNames            ← complete catalogue of all defined names
```

---

## Areas Sensitive to Future Template Version Updates

These are the parts of the code that will break or produce wrong output if the Excel template structure changes.

### 1. Hard-coded Sheet Names

**Location:** `_DATA_SHEET_NAMES` constant, used throughout `_extract_template_data()`

```python
_DATA_SHEET_NAMES = [
    "General Information",
    "Environmental Disclosures",
    "Social Disclosures",
    "Governance Disclosures",
]
```

**Risk:** If any sheet is renamed (e.g. "General Info"), the script will silently skip that sheet entirely — no facts, no sections, no validation rules from it.

---

### 2. Hard-coded Non-Section `template_*` Names

**Location:** `_NON_SECTION_NAMES` set in `_extract_template_data()`

```python
_NON_SECTION_NAMES = {
    "template_currency",
    "template_selected_display_language",
    "template_overall_validation_status",
    "template_starting_date_display",
    "template_translations",
}
```

**Risk:** If new `template_*` non-section names are added (i.e. metadata fields that point to non-bold cells in data sheets), they may be misidentified as section headers (or silently dropped). If existing names are removed, the set becomes stale but harmlessly so.

---

### 3. Section Header Detection via Bold Font

**Location:** Section discovery loop in `_extract_template_data()`

```python
is_bold = cell.font and cell.font.bold
if not is_bold:
    continue  # not a visual section header
```

**Risk:** If the template author removes bold formatting from a header cell, or applies bold formatting to a non-header cell that has a `template_*` name, sections will be missed or wrongly added.

---

### 4. Applicability Flag Detection via String Matching

**Location:** Section and `sectionApplicability` parsing

```python
if "[Always to be reported + If applicable]" in header_text:
    appl = "always+conditional"
elif "[If applicable linked with" in header_text:
    appl = "conditional-linked"
elif "[If applicable]" in header_text:
    appl = "conditional"
elif "[May (optional)]" in header_text:
    appl = "optional"
elif "[Always to be reported]" in header_text:
    appl = "always"
```

**Risk:** These strings are the evaluated text that results from Excel formulas referencing `template_label_always_to_be_reported`, `template_label_if_applicable`, etc. If the English text of those labels changes in the Translations sheet, the matching will fail and all sections will fall back to `applicability: "always"`.

---

### 5. Translations Sheet Layout (Row/Column Coordinates)

**Location:** Translations sheet reading in `_extract_template_data()`

```python
# Row 2 has language codes
for col in range(3, tws.max_column + 1):
    code = tws.cell(row=2, column=col).value
# Rows 4+ have labels
for row in range(4, tws.max_row + 1):
    label_id = tws.cell(row=row, column=1).value
```

**Risk:** The code assumes:
- Language codes are **always in row 2, starting at column 3**
- Label IDs are **always in column 1**
- Data starts at **row 4**

If rows/columns are inserted above or to the left, or if a header row is added, the language codes and/or label IDs will be misread.

---

### 6. Omitted Disclosures Named Range

**Location:** Omitted disclosures extraction

```python
omitted_range_name = (
    "ListOfOmittedDisclosuresDeemedToBeClassifiedOrSensitiveInformation"
)
```

**Risk:** If this defined name is renamed or removed, the omitted disclosures list will always be empty (with a silent fallback). The fallback reads directly from the workbook defined names, so a rename would also defeat the fallback.

---

### 7. Validation Cell Pattern (`template_label_ok` / `template_label_missing_value`)

**Location:** Validation rules extraction (second workbook load)

```python
if "template_label_ok" not in cell.value:
    continue
has_missing = "template_label_missing_value" in formula
```

**Risk:** If the template switches to a different validation indicator pattern (e.g. a different named range or direct text strings), none of the ~997 validation rules will be detected.

---

### 8. Validation Cell Layout (Field on Same Row)

**Location:** `primary_xbrl` detection in validation rules loop

The script assumes the validated XBRL field is **on the same row** as the validation formula cell, searching left then right then by formula cell reference.

**Risk:** If the template is restructured so that validation cells are no longer row-aligned with their data cells, field-to-validation-rule associations will be wrong or missing.

---

### 9. Column C = `template_label_*` Fallback

**Location:** Final fallback for `primary_xbrl` and `field_id`

```python
_label_cell = ws.cell(row=cell.row, column=3)
```

**Risk:** The script assumes column 3 (column C) contains the row's `template_label_*` formula. If the label column moves, the fallback for non-XBRL fields (company name, dates, meta fields) will stop working.

---

### 10. Warning Cell Pattern

**Location:** Warning definitions extraction

```python
_WARNING_RE = re.compile(r"^=?(template_label_\w*warning\w*)$")
```

**Risk:** This regex matches cells containing only a reference to a `template_label_*warning*` name. Multi-cell warning formulas or differently named warning labels would be missed.

---

### 11. `enum_*` Data Validation Formula Convention

**Location:** XBRL → enum mapping via data validation

```python
if not formula or not formula.startswith("enum_"):
    continue
```

**Risk:** This assumes that dropdown data validations always use a bare `enum_*` defined name as their formula (i.e. `formula1 = "enum_BasisForPreparation"`). If any dropdown uses a different formula pattern (e.g. a direct range reference `$A$1:$A$5`), that dropdown will not be linked to its XBRL field.

---

### 12. Side-by-Side Section Column Boundaries

**Location:** Column range calculation for same-row section headers

```python
same_row_headers = [h for h in raw_headers if h["sheet"] == hdr["sheet"] and h["row"] == hdr["row"]]
```

**Risk:** Currently only the GHG / GHG-targets sections are side-by-side (row 16 of Environmental Disclosures). If additional side-by-side sections are introduced, or if existing ones are moved to different rows, the column splitting logic must handle them correctly — which it will automatically only if the new headers also use `template_*` defined names pointing to bold cells.

---

### 13. Yellow Cell Fill Colour Convention

**Location:** Yellow cell scanning in `_extract_template_data()`

```python
_YELLOW_FILL = "FFFFFF99"
if cell.fill.start_color.index == _YELLOW_FILL:
```

**Risk:** The script identifies condition/trigger fields purely by fill colour `FFFFFF99`. If the template author changes the colour convention for these cells, or if openpyxl returns a different colour index format (e.g. theme-based vs. tinted), all 25 yellow cell fields will be missed. Conversely, if unrelated cells are coloured yellow, false positives will be introduced.

The label resolution relies on a `template_label_*` formula in a yellow cell on the same row. If the label formula convention changes, yellow cells may be extracted without meaningful labels.

---

## Summary Table

| # | What It Detects | How | Sensitivity |
|---|----------------|-----|------------|
| 1 | Data sheet names | Hard-coded list | **High** — silent data loss on rename |
| 2 | Non-section template names | Hard-coded set | **Medium** — new metadata names may be misclassified |
| 3 | Section headers | Bold font on `template_*` cell | **Medium** — formatting change drops or adds sections |
| 4 | Applicability flags | English string matching on evaluated text | **High** — label text change breaks all applicability |
| 5 | Translation table layout | Fixed row/column offsets | **High** — row/column insertion breaks i18n |
| 6 | Omitted disclosures field | Hard-coded defined name | **Low** — silent empty list on rename |
| 7 | Validation cells | `template_label_ok` string in formula | **High** — pattern change drops all 101 rules |
| 8 | Field–validation association | Same-row assumption | **Medium** — layout change breaks associations |
| 9 | Non-XBRL field fallback | Column C assumption | **Low** — only affects meta-fields without XBRL names |
| 10 | Warning cells | Regex on `template_label_*warning*` | **Low** — edge cases missed, not a hard failure |
| 11 | Enum dropdowns | `enum_*` formula convention | **Medium** — non-standard dropdowns not linked |
| 12 | Side-by-side sections | Same-row header detection | **Low** — currently only 1 instance |
| 13 | Yellow cell fields | Fill colour `FFFFFF99` | **Medium** — colour change drops all 25 condition fields |

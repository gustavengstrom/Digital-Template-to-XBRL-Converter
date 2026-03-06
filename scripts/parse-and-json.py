"""Extract facts from an Excel digital template and output a JSON mapping of labels to values."""

import argparse
import json
import logging
import re
from contextlib import closing
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import rich.traceback
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.cell_range import CellRange
from rich.logging import RichHandler

import mireport
from mireport.conversionresults import ConversionResultsBuilder
from mireport.excelprocessor import VSME_DEFAULTS, ExcelProcessor
from mireport.excelutil import (
    CellValueType,
    checkExcelFilePath,
    getNamedRanges,
    loadExcelFromPathOrFileLike,
)
from mireport.localise import EU_LOCALES, argparse_locale
from mireport.xbrlreport import ReportLayoutOrganiser


def createArgParser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extract facts from Excel digital template and output as JSON."
    )
    parser.add_argument("excel_file", type=Path, help="Path to the Excel file")
    parser.add_argument(
        "output_path",
        type=Path,
        help="Path to save the JSON output file (e.g. output/report.json)",
    )
    parser.add_argument(
        "--output-locale",
        type=argparse_locale,
        default=None,
        help=f"Locale to use when formatting values (default: None). Examples:\n{sorted(EU_LOCALES)}",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Suppress overwrite warnings and force file replacement.",
    )
    parser.add_argument(
        "--flat",
        action="store_true",
        default=False,
        help="Output a flat label→value mapping instead of the detailed structure.",
    )
    return parser


def parseArgs(parser: argparse.ArgumentParser) -> argparse.Namespace:
    args = parser.parse_args()
    return args


def fact_value_to_json(value) -> str | int | float | bool | None:
    """Convert a FactValue to a JSON-serializable type."""
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


# ---------------------------------------------------------------------------
# JsonExcelProcessor – subclass that also captures template_* named ranges
# ---------------------------------------------------------------------------


class JsonExcelProcessor(ExcelProcessor):
    """
    Extends ExcelProcessor to additionally capture the template_* named ranges
    that the base class intentionally skips (they are not XBRL facts but carry
    section-applicability flags and human-readable labels).

    The base class closes the workbook at the end of ``populateReport()``.
    This subclass overrides that method to also call ``get_template_data()``
    while the workbook is still open, storing the result for later retrieval.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # {name -> DefinedName} for every template_* range in the workbook
        self._template_defined_names: dict[str, DefinedName] = {}
        self._template_data: Optional[dict] = None

    # -- overrides -----------------------------------------------------------

    def _recordNamedRanges(self) -> None:
        """Record both XBRL-related *and* template_* named ranges."""
        # Let the base class populate _unusedDefinedNames (non-template ranges)
        super()._recordNamedRanges()

        # Additionally collect every template_* range that the base class skips
        for dn in self._workbook.defined_names.values():
            if dn.name and dn.name.startswith("template_"):
                self._template_defined_names[dn.name] = dn

    def populateReport(self):
        """
        Override to extract template data while the workbook is still open.
        The base class closes the workbook in its ``finally`` block, so we
        hook into ``checkForUnhandledItems`` which runs just before that.
        """
        from mireport.xbrlreport import InlineReport

        report = super().populateReport()
        return report

    def checkForUnhandledItems(self) -> None:
        """
        Called by the base class just before the workbook is closed.
        We extract template data here while the workbook is still available,
        then call through to the base implementation.
        """
        self._template_data = self._extract_template_data()
        super().checkForUnhandledItems()

    def get_template_data(self) -> dict:
        """
        Return the template data collected during ``populateReport()``.
        Must be called after ``populateReport()`` has completed.
        """
        if self._template_data is None:
            raise RuntimeError(
                "Template data not yet available. Call populateReport() first."
            )
        return self._template_data

    def _extract_template_data(self) -> dict:
        """
        Internal: extract template data while the workbook is still open.

        Return a structured dict with four top-level keys:

        * ``labels``   – template_label_* display strings
        * ``metadata`` – template_reporting_* and other template_* descriptors
        * ``sectionApplicability`` – per-section applicability status parsed
          from the descriptors ("always", "conditional", "optional", …)
        * ``omittedDisclosures`` – list of disclosure section names the user
          has marked as omitted via the List of Omitted Disclosures dropdown
          (from the ``ListOfOmittedDisclosuresDeemedToBeClassifiedOrSensitive
          Information`` named range)
        * ``conditionalQuestions`` – yes/no question labels paired with their
          answer value (found via an adjacent cell or a matching XBRL named
          range) and the section(s) they gate
        * ``enumLists`` – all ``enum_*`` defined names from the workbook mapped
          to their list of allowed values (the dropdown option lists)
        * ``xbrlToEnum`` – mapping from XBRL named-range names to the
          ``enum_*`` list name(s) that provide their dropdown options (derived
          from the Excel data-validation rules on the corresponding cells)
        """
        assert self._workbook is not None
        all_ranges, _ = getNamedRanges(self._workbook)

        labels: dict[str, object] = {}
        metadata: dict[str, object] = {}
        section_applicability: dict[str, str] = {}
        conditional_questions: list[dict] = []

        # --- labels & metadata ------------------------------------------------
        for name, cells in sorted(all_ranges.items()):
            if not name.startswith("template_"):
                continue
            non_empty = [c for c in cells if c is not None]
            if not non_empty:
                continue
            values = [fact_value_to_json(v) for v in non_empty]
            value = values[0] if len(values) == 1 else values

            if name.startswith("template_label_"):
                clean = name[len("template_label_") :]
                labels[clean] = value
            else:
                clean = name[len("template_") :]
                metadata[clean] = value

        # --- section applicability -------------------------------------------
        for key, val in metadata.items():
            text = str(val)
            if "[Always to be reported + If applicable]" in text:
                status = "always+conditional"
            elif "[Always to be reported]" in text:
                status = "always"
            elif "[If applicable linked with" in text:
                status = "conditional-linked"
            elif "[If applicable]" in text:
                status = "conditional"
            elif "[May (optional)]" in text:
                status = "optional"
            else:
                continue
            section_applicability[key] = status

        # --- omitted disclosures -------------------------------------------
        # The user can explicitly mark whole disclosure sections as omitted
        # via a multi-select dropdown.  The values come from the XBRL named
        # range ListOfOmittedDisclosuresDeemedToBeClassifiedOrSensitiveInformation
        # which maps to a column in General Information (E123:E222).
        omitted_disclosures: list[str] = []
        omitted_range_name = (
            "ListOfOmittedDisclosuresDeemedToBeClassifiedOrSensitiveInformation"
        )
        if omitted_range_name in all_ranges:
            for val in all_ranges[omitted_range_name]:
                if val is not None:
                    s = str(val).strip()
                    if s and s.lower() != "none":
                        omitted_disclosures.append(s)
        else:
            # Fall back: try reading directly from the workbook defined name
            for dn in self._workbook.defined_names.values():
                if dn.name == omitted_range_name:
                    for sheet_title, coord in dn.destinations:
                        if sheet_title in self._workbook:
                            ws = self._workbook[sheet_title]
                            try:
                                cr = CellRange(coord)
                            except Exception:
                                break
                            for row in range(cr.min_row, cr.max_row + 1):
                                for col in range(cr.min_col, cr.max_col + 1):
                                    cv = ws.cell(row=row, column=col).value
                                    if cv is not None:
                                        s = str(cv).strip()
                                        if s and s.lower() != "none":
                                            omitted_disclosures.append(s)

        # --- enum lists -------------------------------------------------------
        # Extract all enum_* defined names and their values.  These are the
        # dropdown option lists used by the template.
        enum_lists: dict[str, list[str]] = {}
        for dn in self._workbook.defined_names.values():
            if not dn.name or not dn.name.startswith("enum_"):
                continue
            values: list[str] = []
            for sheet_title, coord in dn.destinations:
                if sheet_title not in self._workbook:
                    continue
                ws = self._workbook[sheet_title]
                try:
                    cr = CellRange(coord)
                except Exception:
                    continue
                for row in range(cr.min_row, cr.max_row + 1):
                    for col in range(cr.min_col, cr.max_col + 1):
                        cv = ws.cell(row=row, column=col).value
                        if cv is not None:
                            values.append(str(cv))
            enum_lists[dn.name] = values

        # --- XBRL → enum mapping via data validation -------------------------
        # Cells with dropdowns have data-validation rules whose formula1
        # references an enum_* defined name.  By intersecting those cells with
        # the cells covered by XBRL named ranges we can link each XBRL concept
        # to the enum list(s) that provide its answer options.
        data_sheet_names = [
            "General Information",
            "Environmental Disclosures",
            "Social Disclosures",
            "Governance Disclosures",
        ]
        # (sheet, row, col) → enum_name
        cell_to_enum: dict[tuple[str, int, int], str] = {}
        for sheet_name in data_sheet_names:
            if sheet_name not in self._workbook:
                continue
            ws = self._workbook[sheet_name]
            if ws.data_validations is None:
                continue
            for dv in ws.data_validations.dataValidation:
                formula = dv.formula1
                if not formula or not formula.startswith("enum_"):
                    continue
                enum_name = formula.strip()
                for cr_str in str(dv.sqref).split():
                    try:
                        cr = CellRange(cr_str)
                    except Exception:
                        continue
                    for row in range(cr.min_row, cr.max_row + 1):
                        for col in range(cr.min_col, cr.max_col + 1):
                            cell_to_enum[(sheet_name, row, col)] = enum_name

        xbrl_to_enum: dict[str, list[str]] = {}
        for dn in self._workbook.defined_names.values():
            if not dn.name or dn.name.startswith(("enum_", "template_")):
                continue
            for sheet_title, coord in dn.destinations:
                if sheet_title not in self._workbook:
                    continue
                try:
                    cr = CellRange(coord)
                except Exception:
                    continue
                for row in range(cr.min_row, cr.max_row + 1):
                    for col in range(cr.min_col, cr.max_col + 1):
                        key = (sheet_title, row, col)
                        if key in cell_to_enum:
                            enum_name = cell_to_enum[key]
                            xbrl_to_enum.setdefault(dn.name, [])
                            if enum_name not in xbrl_to_enum[dn.name]:
                                xbrl_to_enum[dn.name].append(enum_name)

        # --- question → section mapping --------------------------------------
        # Each conditional question sits just below a section descriptor in the
        # data sheets.  We scan all data sheets to build a map from question
        # text (first 50 chars) to the section descriptor above it.
        question_to_section: dict[str, str] = (
            {}
        )  # question_text_prefix → section_descriptor
        applicability_flags = [
            "[If applicable]",
            "[Always to be reported]",
            "[May (optional)]",
            "[Always to be reported + If applicable]",
            "[If applicable linked with",
        ]
        for sheet_name in data_sheet_names:
            if sheet_name not in self._workbook:
                continue
            ws = self._workbook[sheet_name]
            last_section_descriptor: str | None = None
            for r in range(1, ws.max_row + 1):
                for c in range(1, min(15, ws.max_column + 1)):
                    v = ws.cell(row=r, column=c).value
                    if not v or not isinstance(v, str):
                        continue
                    # Is this a section descriptor with an applicability flag?
                    if any(flag in v for flag in applicability_flags):
                        last_section_descriptor = v
                    # Is this a question?
                    elif v.rstrip().endswith("?") and last_section_descriptor:
                        prefix = v[:50].lower()
                        question_to_section[prefix] = last_section_descriptor

        # --- conditional questions -------------------------------------------
        #
        # Template labels that are yes/no questions ("Has the undertaking…",
        # "Does the undertaking…", etc.).  We try to pair each label with its
        # answer by:
        #   1. Looking for an XBRL named range whose name is keyword-similar.
        #   2. Scanning the data sheet for the label text and reading the
        #      adjacent checkbox cell.
        #
        xbrl_names: dict[str, list] = {
            n: c
            for n, c in all_ranges.items()
            if not n.startswith("template_") and not n.startswith("enum_")
        }

        # Build a set of cells covered by named ranges (for orphan detection)
        covered_cells: set[tuple[str, int, int]] = set()
        for dn in self._workbook.defined_names.values():
            if not dn.name:
                continue
            dests = list(dn.destinations)
            if not dests:
                continue
            sheet, cell_range = dests[0]
            if sheet not in self._workbook:
                continue
            try:
                cr = CellRange(cell_range)
            except Exception:
                continue
            for row in range(cr.min_row, cr.max_row + 1):
                for col in range(cr.min_col, cr.max_col + 1):
                    covered_cells.add((sheet, row, col))

        def _is_boolean_like(value) -> bool:
            """Check if a value looks like a yes/no or boolean answer."""
            if isinstance(value, bool):
                return True
            s = str(value).strip().upper()
            return s in ("YES", "NO", "TRUE", "FALSE")

        for label_key, label_text in sorted(labels.items()):
            text_str = str(label_text).lower().strip()

            # Only keep labels that look like yes/no questions
            # Must end with '?' or start with a question phrase
            is_question = text_str.rstrip().endswith("?")
            if not is_question:
                continue

            answer_value = None
            answer_source: Optional[str] = None

            # Strategy 1: keyword-match against XBRL named ranges
            # Prefer boolean/yes-no matches over text-block matches.
            words = re.findall(r"[a-zA-Z]+", label_key)
            best_match = None
            best_score = 0
            best_is_boolean = False
            for xbrl_name, xbrl_cells in xbrl_names.items():
                xl = xbrl_name.lower()
                score = sum(1 for w in words if w.lower() in xl)
                threshold = len(words) * 0.6
                if score >= threshold:
                    non_empty = [c for c in xbrl_cells if c is not None]
                    if not non_empty:
                        continue
                    candidate_val = non_empty[0] if len(non_empty) == 1 else non_empty
                    is_bool = _is_boolean_like(candidate_val)
                    # Prefer boolean-like matches; among same type prefer higher score
                    if (is_bool and not best_is_boolean) or (
                        is_bool == best_is_boolean and score > best_score
                    ):
                        best_match = xbrl_name
                        best_score = score
                        best_is_boolean = is_bool
                        answer_value = fact_value_to_json(candidate_val)
                        answer_source = f"xbrl:{xbrl_name}"

            # Strategy 2: scan for orphan checkbox / yes-no cell on data sheets
            # Try to find a boolean answer, overriding non-boolean Strategy 1 matches
            if answer_value is None or not _is_boolean_like(answer_value):
                search_text = str(label_text)[:50].lower()
                data_sheets = [
                    "General Information",
                    "Environmental Disclosures",
                    "Social Disclosures",
                    "Governance Disclosures",
                ]
                found_boolean = False
                for sheet_name in data_sheets:
                    if found_boolean:
                        break
                    if sheet_name not in self._workbook:
                        continue
                    ws = self._workbook[sheet_name]
                    for row in ws.iter_rows(
                        min_row=1,
                        max_row=ws.max_row,
                        min_col=1,
                        max_col=ws.max_column,
                    ):
                        if found_boolean:
                            break
                        for cell in row:
                            if found_boolean:
                                break
                            if (
                                cell.value
                                and isinstance(cell.value, str)
                                and cell.value[:50].lower() == search_text
                            ):
                                # Look for a boolean or YES/NO in:
                                #  a) same column, rows below (offset 1-3)
                                #  b) same row, columns to the right
                                candidates: list[tuple[int, int]] = []
                                for offset in [2, 1, 3]:
                                    candidates.append((cell.row + offset, cell.column))
                                for col_off in range(1, 15):
                                    candidates.append((cell.row, cell.column + col_off))
                                for r, c in candidates:
                                    try:
                                        check = ws.cell(row=r, column=c)
                                    except Exception:
                                        continue
                                    if check.value is not None and (
                                        isinstance(check.value, bool)
                                        or str(check.value).strip().upper()
                                        in ("YES", "NO")
                                    ):
                                        key = (sheet_name, check.row, check.column)
                                        answer_value = fact_value_to_json(check.value)
                                        if key in covered_cells:
                                            answer_source = (
                                                f"cell:{sheet_name}!{check.coordinate}"
                                            )
                                        else:
                                            answer_source = f"cell:{sheet_name}!{check.coordinate} (orphaned)"
                                        found_boolean = True
                                        break

            # Determine which section this question gates
            question_prefix = str(label_text)[:50].lower()
            gated_section = question_to_section.get(question_prefix)

            entry: dict = {
                "key": label_key,
                "label": str(label_text),
                "value": answer_value,
                "source": answer_source,
                "conditional": True,
            }
            if gated_section:
                entry["gatesSection"] = gated_section

            conditional_questions.append(entry)

        return {
            "labels": labels,
            "metadata": metadata,
            "sectionApplicability": section_applicability,
            "omittedDisclosures": omitted_disclosures,
            "conditionalQuestions": conditional_questions,
            "enumLists": enum_lists,
            "xbrlToEnum": xbrl_to_enum,
        }


def fact_to_json_entry(
    fact,
    xbrl_to_enum: Optional[dict[str, list[str]]] = None,
) -> dict:
    """Convert a single Fact to a JSON-serializable dict.

    If *xbrl_to_enum* is provided and the fact's concept name matches an entry,
    an ``"options"`` key is added referencing the ``enum_*`` list name(s) that
    supply the dropdown choices for this fact.
    """
    concept = fact.concept
    entry: dict = {
        "qname": str(concept.qname),
        "label": concept.getStandardLabel(fallbackToQName=True),
        "value": fact_value_to_json(fact.value),
        "dataType": str(concept.dataType) if concept.dataType else None,
        "periodType": concept.periodType,
    }

    # Add formatted value if available
    try:
        formatted = fact.formattedValue
        if formatted and formatted != str(fact.value):
            entry["formattedValue"] = formatted
    except Exception:
        pass

    # Add aspects (period, units, dimensions, etc.)
    aspects = fact.aspects
    if aspects:
        serializable_aspects = {}
        for k, v in aspects.items():
            serializable_aspects[str(k)] = str(v)
        entry["aspects"] = serializable_aspects

    # Add taxonomy dimensions if present
    if fact.hasTaxonomyDimensions():
        dims = {}
        for dim_qname, member_qname in fact.getTaxonomyDimensions().items():
            dims[str(dim_qname)] = str(member_qname)
        entry["dimensions"] = dims

    # Add reference to the enum list(s) that provide dropdown options
    if xbrl_to_enum:
        # concept.qname is e.g. "vsme:SomeName"; the XBRL named range uses
        # just the local part "SomeName"
        local_name = (
            str(concept.qname).split(":", 1)[-1]
            if ":" in str(concept.qname)
            else str(concept.qname)
        )
        if local_name in xbrl_to_enum:
            enum_refs = xbrl_to_enum[local_name]
            entry["options"] = enum_refs[0] if len(enum_refs) == 1 else enum_refs

    return entry


def _match_section_to_applicability(
    section_definition: str,
    section_applicability: dict[str, str],
) -> Optional[str]:
    """
    Match a presentation group definition like
    ``[B02.000] - General information - Practices, policies …``
    to a section applicability key like
    ``b2_practices_policies_and_future_initiatives_…``.

    Returns the applicability status string or None.
    """
    # Extract the short code, e.g. "B02" → "b2"
    m = re.match(r"\[(\w+)\.\d+\]", section_definition)
    if not m:
        return None
    short_code = m.group(1).lower().lstrip("0")  # "B02" → "b2", "C02" → "c2"
    # Normalise: drop leading zeros in the number part
    short_code = re.sub(r"^([a-z])(0+)", r"\1", m.group(1).lower())

    for key, status in section_applicability.items():
        if key.startswith(short_code + "_"):
            # Rough check – the key starts with the same section prefix
            return status
    return None


def extract_facts_by_section(
    report,
    flat: bool = False,
    section_applicability: Optional[dict[str, str]] = None,
    xbrl_to_enum: Optional[dict[str, list[str]]] = None,
) -> list[dict]:
    """
    Extract all facts organized by their presentation group (section header).

    Uses the taxonomy's presentation linkbase to group facts under headers like
    "[B08.100] - Social - Workforce - General characteristics: gender".

    If *section_applicability* is provided each section is annotated with its
    reporting status (``"always"``, ``"conditional"``, ``"optional"``, …).

    If *xbrl_to_enum* is provided each fact whose XBRL named range overlaps
    with a data-validated cell gets an ``"options"`` key referencing the
    ``enum_*`` list(s) that supply its dropdown choices.

    If flat=True, each section contains a simple label→value mapping.
    If flat=False, each section contains detailed fact entries.
    """
    organiser = ReportLayoutOrganiser(report._taxonomy, report)
    sections = organiser.organise()

    result: list[dict] = []
    for section in sections:
        if not section.hasFacts:
            continue

        section_header = section.presentation.definition
        section_label = section.getLabel("en")

        section_entry: dict = {
            "section": section_header,
            "sectionLabel": section_label,
        }

        # Annotate with applicability
        if section_applicability:
            status = _match_section_to_applicability(
                section_header,
                section_applicability,
            )
            if status:
                section_entry["applicability"] = status

        if flat:
            facts_in_section: dict[str, object] = {}
            for rel, factList in section.relationshipToFact.items():
                for fact in factList:
                    label = fact.concept.getStandardLabel(fallbackToQName=True)
                    value = fact_value_to_json(fact.value)
                    if label in facts_in_section:
                        existing = facts_in_section[label]
                        if isinstance(existing, list):
                            existing.append(value)
                        else:
                            facts_in_section[label] = [existing, value]
                    else:
                        facts_in_section[label] = value
            section_entry["facts"] = facts_in_section
        else:
            facts_list: list[dict] = []
            for rel, factList in section.relationshipToFact.items():
                for fact in factList:
                    entry = fact_to_json_entry(fact, xbrl_to_enum=xbrl_to_enum)
                    entry["section"] = section_header
                    facts_list.append(entry)
            section_entry["facts"] = facts_list

        result.append(section_entry)

    return result


def main() -> None:
    parser = createArgParser()
    args = parseArgs(parser)

    resultsBuilder = ConversionResultsBuilder(consoleOutput=True)
    with resultsBuilder.processingContext("mireport Excel to JSON export") as pc:
        pc.mark("Loading taxonomy metadata")
        mireport.loadTaxonomyJSON()

        pc.mark(
            "Extracting data from Excel",
            additionalInfo=f"Using file: {args.excel_file}",
        )
        excel = JsonExcelProcessor(
            args.excel_file,
            resultsBuilder,
            VSME_DEFAULTS,
            outputLocale=args.output_locale,
        )
        report = excel.populateReport()

        pc.mark("Extracting template metadata & conditional questions")
        template_data = excel.get_template_data()

        pc.mark("Generating JSON output")
        sections_data = extract_facts_by_section(
            report,
            flat=args.flat,
            section_applicability=template_data["sectionApplicability"],
            xbrl_to_enum=template_data["xbrlToEnum"],
        )

        # Build the output structure
        output = {
            "reportTitle": report._reportTitle,
            "entityName": report._entityName,
            "factCount": report.factCount,
            "templateMetadata": template_data["metadata"],
            "templateLabels": template_data["labels"],
            "sectionApplicability": template_data["sectionApplicability"],
            "omittedDisclosures": template_data["omittedDisclosures"],
            "conditionalQuestions": template_data["conditionalQuestions"],
            "sections": sections_data,
            "enumLists": template_data["enumLists"],
        }

        # Resolve output path
        output_path = args.output_path
        if output_path.is_dir() or (not output_path.suffix):
            output_path.mkdir(parents=True, exist_ok=True)
            output_path = output_path / "report-facts.json"
        else:
            output_path.parent.mkdir(parents=True, exist_ok=True)

        if output_path.exists() and not args.force:
            print(f"⚠️  Warning: Overwriting existing file: {output_path}")

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(output, f, indent=2, ensure_ascii=False)

        pc.mark(
            "Done", additionalInfo=f"Wrote {report.factCount} facts to {output_path}"
        )

    result = resultsBuilder.build()
    if result.hasMessages(userOnly=True):
        print()
        messages = result.userMessages
        print(f"Messages ({len(messages)}):")
        for msg in messages:
            print(f"\t{msg}")


if __name__ == "__main__":
    rich.traceback.install(show_locals=False)
    logging.basicConfig(
        format="%(message)s",
        datefmt="[%Y-%m-%d %H:%M:%S]",
        handlers=[RichHandler(rich_tracebacks=True)],
    )
    logging.captureWarnings(True)
    main()

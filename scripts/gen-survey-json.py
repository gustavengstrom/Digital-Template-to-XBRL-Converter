#!/usr/bin/env python3
"""
gen-survey-json.py — Convert report-facts.json into survey_data JSON files.

Reads the structured output of parse-and-json.py and converts each section into
a flat list of question objects following the survey_data_proxy format described
in .github/instructions/survey_data.instructions.md.

Usage:
    python scripts/gen-survey-json.py [INPUT_JSON] [OUTPUT_DIR]

Defaults:
    INPUT_JSON = output/report-facts.json
    OUTPUT_DIR = output/surveys
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SURVEY_PREFIX = "VSME"

# Mapping from report-facts inputType → survey answer_type
_ANSWER_TYPE_MAP: dict[str, str] = {
    "boolean": "RADIO",
    "text": "OPEN",
    "textarea": "OPEN_LARGE",
    "number": "NUMERIC",
    "monetary": "NUMERIC",
    "percent": "NUMERIC",
    "select": "DROPDOWN",
    "dropdown": "DROPDOWN",
    "multiselect": "CHECKBOX",
    "date": "DATE",
    "year": "DROPDOWN",
}

# Synthetic fields for orphaned column labels.
# When a section's columnLabels contains a column that has no
# corresponding data field (the parser didn't emit one because
# the column holds a dimension axis value rather than an XBRL
# data point), we inject a synthetic field so the survey has a
# question for that column.
#
# Each entry maps ``(section_id, column_label_key)`` to the
# synthetic field definition that will be inserted at the
# beginning of the table fields.
_SYNTHETIC_COLUMNS: dict[tuple[str, str], dict[str, Any]] = {
    ("b4_pollution_of_air_water_and_soil", "pollutant"): {
        "fieldId": "synthetic_pollutant",
        "source": "synthetic",
        "labelKey": "pollutant",
        "templateLabelKey": "template_label_pollutant",
        "label": "Pollutant",
        "dataType": "enumeration",
        "inputType": "dropdown",
        "isRequired": True,
        "options": ["enum_ListPollutants"],
    },
    ("b7_waste_generated", "type_of_waste"): {
        "fieldId": "synthetic_type_of_waste",
        "source": "synthetic",
        "labelKey": "type_of_waste",
        "templateLabelKey": "template_label_type_of_waste",
        "label": "Type of waste",
        "dataType": "enumeration",
        "inputType": "dropdown",
        "isRequired": True,
        "options": ["enum_ListTypeOfWastes"],
    },
}


# Sections whose fields represent repeatable rows (dynamic groups).
# Each entry maps a sectionId to its dynamic-group configuration.
_DYNAMIC_SECTIONS: dict[str, dict[str, Any]] = {
    "b1_list_of_subsidiaries": {
        "dynamic_header": "Subsidiary",
        "dynamic_max_items": None,  # unlimited
        "tag_field_qname": "vsme:NameOfTheSubsidiary",
    },
    "b1_list_of_sites": {
        "dynamic_header": "Site",
        "dynamic_max_items": None,
        "tag_field_qname": "vsme:AddressOfSite",
    },
    "b4_pollution_of_air_water_and_soil": {
        "dynamic_header": "Pollutant",
        "dynamic_max_items": None,
        "tag_field_qname": None,
        "tag_field_key": "synthetic_pollutant",
        "table_fields_only": True,
    },
    "b5_sites_in_biodiversity_sensitive_areas": {
        "dynamic_header": "Site",
        "dynamic_max_items": None,
        "tag_field_qname": None,
        "table_fields_only": True,
    },
    "b7_waste_generated": {
        "dynamic_header": "Waste type",
        "dynamic_max_items": None,
        "tag_field_qname": None,
        "tag_field_key": "synthetic_type_of_waste",
        "table_fields_only": True,
    },
    "b7_annual_mass_flow_of_relevant_materials_used": {
        "dynamic_header": "Material",
        "dynamic_max_items": None,
        "tag_field_qname": "vsme:NameOfMaterialUsed",
        "table_fields_only": True,
    },
    "b8_workforce_general_characteristics_country_of_employment": {
        "dynamic_header": "Country",
        "dynamic_max_items": None,
        "tag_field_qname": None,
        "table_fields_only": True,
    },
}


# Section-specific field exclusion rules.
# Each entry maps a sectionId to a predicate function that returns True
# for fields that should be *excluded* from survey output.
#
# b1_list_of_sites: The yellow cells at H447:H449 contain an OpenStreetMaps
# automatic-geolocation checkbox (template_label_automatic_geolocation).
# This is an Excel UI feature, not a survey question, so we exclude it.
_EXCLUDED_FIELDS: dict[str, Any] = {
    "b1_list_of_sites": {
        "templateLabelKeys": {"template_label_automatic_geolocation"},
    },
}


_ABBREVIATIONS: dict[str, str] = {
    "information_on_the_report_necessary_for_XBRL": "info_xbrl",
    "information_on_previous_reporting_period": "info_prev_period",
    "b1_basis_for_preparation_and_other_undertakings_general_information": "b1_basis_prep",
    "b1_list_of_subsidiaries": "b1_subsidiaries",
    "b1_disclosure_of_sustainability_related_certifications_or_labels": "b1_certifications",
    "b1_list_of_sites": "b1_sites",
    "b2_practices_policies_and_future_initiatives_for_transitioning_towards_a_more_sustainable_economy": "b2_practices_policies",
    "b2_cooperative_specific_disclosures": "b2_coop_disclosures",
    "c2_description_of_practices_policies_and_future_initiatives_for_transitioning_towards_a_more_sustainable_economy": "c2_practices_policies",
    "c1_strategy_business_model_and_sustainability": "c1_strategy",
    "c1_strategy_business_model_and_sustainability_if_applicable": "c1_strategy_if_applicable",
    "disclosure_of_any_other_general_and_or_entity_specific_information_on_the_reporting_period": "other_general_info",
    "b3_total_energy_consumption_in_MWh": "b3_total_energy",
    "b3_breakdown_of_energy_consumption_in_MWh": "b3_energy_breakdown",
    "b3_estimated_greenhouse_gas_emissions_considering_the_GHG_protocol_version_2004_in_tCO2e": "b3_estimated_ghg_emissions",
    "c3_GHG_reduction_targets_in_tC02e": "c3_ghg_targets",
    "c3_disclosure_of_list_of_main_actions_the_entity_seeks_in_order_to_achieve_its_targets": "c3_main_actions",
    "c3_transition_plan_for_undertakings_operating_in_high_climate_impact_sectors": "c3_transition_plan",
    "b3_greenhouse_gas_emission_intensity_per_turnover": "b3_ghg_intensity",
    "b4_pollution_of_air_water_and_soil": "b4_pollution",
    "b5_sites_in_biodiversity_sensitive_areas": "b5_biodiversity_sites",
    "b5_biodiversity_land_use": "b5_land_use",
    "b6_water_withdrawal": "b6_water_withdrawal",
    "b6_water_consumption": "b6_water_consumption",
    "b7_description_of_circular_economy_principles": "b7_circular_economy",
    "b7_waste_generated": "b7_waste",
    "b7_annual_mass_flow_of_relevant_materials_used": "b7_mass_flow",
    "c4_climate_risks": "c4_climate_risks",
    "disclosure_of_any_other_environmental_and_or_entity_specific_enviromental_disclosures": "other_environmental",
    "b8_workforce_general_characteristics_type_of_contract": "b8_workforce_contract",
    "b8_workforce_general_characteristics_gender": "b8_workforce_gender",
    "b8_workforce_general_characteristics_country_of_employment": "b8_workforce_country",
    "b8_workforce_general_characteristics_turnover_rate": "b8_workforce_turnover",
    "b9_workforce_health_and_safety": "b9_health_safety",
    "b10_workforce_remuneration_collective_bargaining_and_training": "b10_remuneration_training",
    "b10_workforce_remuneration_collective_bargaining_and_training_always_reported": "b10_remuneration_training_always",
    "c5_additional_general_workforce_characteristics": "c5_additional_workforce",
    "c6_additional_own_workforce_information_human_rights_policies_and_processes": "c6_human_rights",
    "c7_severe_negative_human_rights_incidents": "c7_negative_hr_incidents",
    "disclosure_of_any_other_social_and_or_entity_specific_social_disclosures": "other_social",
    "b11_convictions_and_fines_for_corruption_and_bribery": "b11_corruption_fines",
    "c8_revenues_from_certain_sectors": "c8_sector_revenues",
    "c8_exclusion_from_EU_reference_benchmarks": "c8_eu_benchmarks",
    "c9_gender_diversity_ratio_in_the_governance_body": "c9_gender_diversity",
    "disclosure_of_any_other_governance_and_or_entity_specific_governance_disclosures": "other_governance",
}


def _should_exclude_field(section_id: str, field: dict) -> bool:
    """Return True if *field* should be excluded from survey output."""
    # Section-specific exclusion rules
    rule = _EXCLUDED_FIELDS.get(section_id)
    if rule:
        excluded_keys = rule.get("templateLabelKeys", set())
        # Check field-level templateLabelKey (yellow cells)
        if field.get("templateLabelKey") in excluded_keys:
            return True

    # Display-only unit fields — unit selection cells whose value is a
    # formula-computed label (e.g. =template_label_kilograms) with no
    # selectable options.  These are not editable in a survey context.
    if field.get("source") == "unitSelection" and not field.get("options"):
        return True

    return False


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _short_id(value: str) -> str:
    """Produce a stable 5-char alphanumeric hash from *value*."""
    return hashlib.sha256(value.encode()).hexdigest()[:5]


def _make_question_id(section_id: str, field: dict, field_index: int = 0) -> str:
    """
    Build a unique question ID.

    Format: {survey_abbrev}_{field_short}
    where survey_abbrev is the abbreviated section name from _ABBREVIATIONS
    (falling back to SURVEY_PREFIX if not found) and field_short is a stable
    5-char alphanumeric hash derived from the section, field identity, and
    field index.
    """
    prefix = _ABBREVIATIONS.get(section_id, SURVEY_PREFIX)

    field_key = (
        field.get("qname") or field.get("fieldId") or field.get("labelKey", "unknown")
    )
    # Include column label key to disambiguate table cells sharing the same
    # qname but in different columns (e.g. renewable vs non-renewable).
    col_key = ""
    if field.get("columnLabel"):
        col_key = field["columnLabel"].get("key", "")
    # Include row label key for extra disambiguation in matrix-style tables
    row_key = ""
    if field.get("rowLabel"):
        row_key = field["rowLabel"].get("key", "")

    # Always include the field index in the hash to produce stable, unique IDs.
    composite = f"{section_id}::{field_key}::{col_key}::{row_key}::{field_index}"
    return f"{prefix}_{_short_id(composite)}"


def _is_instructional_text(text: str) -> bool:
    """Return True if the text looks like instructional/helper text, not a label."""
    if not text:
        return False
    return text.startswith("ℹ️") or len(text) > 200


def _resolve_template_label(
    template_label_key: str, defined_names: dict[str, dict]
) -> str:
    """Resolve a templateLabelKey to its display value via definedNames."""
    if not template_label_key:
        return ""
    dn = defined_names.get(template_label_key)
    if dn and dn.get("value"):
        return str(dn["value"])
    # Fallback: return the key itself
    return template_label_key


def _resolve_unit_field_data_label(field: dict, section: dict) -> str:
    """For a unitSelection field, find the sibling data field's label.

    Unit selection fields have ``fieldId`` like ``AmountOfEmissionToAir_unit``.
    The sibling XBRL data field has ``qname == "vsme:AmountOfEmissionToAir"``.
    Returns the sibling's XBRL label (e.g. "Amount of emission to air"), or
    the humanised concept name as a fallback.
    """
    field_id = field.get("fieldId", "")
    if not field_id.endswith("_unit"):
        return ""
    concept = field_id[: -len("_unit")]  # e.g. "AmountOfEmissionToAir"

    # Search sibling fields for the matching XBRL data field
    for sibling in section.get("fields", []):
        qname = sibling.get("qname", "")
        if qname and qname.split(":", 1)[-1] == concept:
            return sibling.get("label", "") or concept
    # Fallback: humanise the concept name (CamelCase → spaced)
    import re as _re

    return _re.sub(r"(?<=[a-z])(?=[A-Z])", " ", concept)


def _build_question_name(
    field: dict, section: dict, defined_names: dict[str, dict]
) -> str:
    """
    Build the question name by resolving templateLabelKey values to their
    display text (looked up from definedNames).

    Rules:
    - Unit selection fields: use the section-wide instruction label
      disambiguated with the sibling data field's label in parentheses.
    - If both rowLabel and columnLabel exist: use the resolved rowLabel
      display value followed by the resolved columnLabel display value
      in parentheses.
    - If only rowLabel exists: use its resolved display value.
    - If only columnLabel exists: use its resolved display value.
    - Otherwise fall back to the field-level templateLabelKey,
      then labelKey, then XBRL label.
    """
    # Unit selection fields: disambiguate with the sibling data field's label
    if field.get("source") == "unitSelection":
        base_label = field.get("label", "")
        base_usable = base_label and not _is_instructional_text(base_label)
        # Consolidated unit fields (multiple _unit fields sharing one Excel
        # cell) don't need data-label disambiguation — the section-wide
        # instruction label is sufficient on its own.
        if field.get("unitFieldIds") and base_usable:
            return base_label

        data_label = _resolve_unit_field_data_label(field, section)
        # If the base label is instructional/unhelpful, use just the data label
        if base_usable and data_label:
            # Avoid redundancy when the labels substantially overlap
            bl = base_label.lower().strip()
            dl = data_label.lower().strip()
            if dl in bl or bl in dl:
                return f"Unit selection — {data_label}"
            return f"{base_label} — {data_label}"
        if data_label:
            return f"Unit selection — {data_label}"

    row_label = field.get("rowLabel")
    col_label = field.get("columnLabel")

    row_tlk = row_label.get("templateLabelKey", "") if row_label else ""
    col_tlk = col_label.get("templateLabelKey", "") if col_label else ""

    row_text = _resolve_template_label(row_tlk, defined_names)
    col_text = _resolve_template_label(col_tlk, defined_names)

    # Skip instructional row labels (e.g. long OpenStreetMaps disclaimer in
    # b1_list_of_sites) — fall through to XBRL label instead.
    row_usable = row_tlk and not _is_instructional_text(row_text)
    col_usable = col_tlk and not _is_instructional_text(col_text)

    # For XBRL fields whose rowLabel was assigned by the "nearest row above"
    # fallback, the rowLabel key may not semantically match the field.
    # When the field has its own labelKey that differs from the rowLabel key
    # AND resolves to a proper template_label display name, prefer the
    # field's own label (e.g. "Gender diversity ratio in governance body"
    # instead of "Number of male board members…").
    if (
        row_usable
        and field.get("qname")
        and field.get("labelKey")
        and row_label
        and row_label.get("key") != field.get("labelKey")
    ):
        own_label_key = f"template_label_{field['labelKey']}"
        own_label_text = _resolve_template_label(own_label_key, defined_names)
        # Only skip the rowLabel if the field's own label resolves to a
        # non-fallback, non-instructional value (i.e. the defined name
        # exists and the resolution didn't just echo the key back).
        if (
            own_label_text
            and own_label_text != own_label_key
            and not _is_instructional_text(own_label_text)
        ):
            row_usable = False

    if row_usable and col_usable:
        return f"{row_text} ({col_text})"
    if row_usable:
        return row_text
    if col_usable:
        return col_text

    # Field-level templateLabelKey (yellow cells)
    field_tlk = field.get("templateLabelKey", "")
    if field_tlk:
        field_text = _resolve_template_label(field_tlk, defined_names)
        if not _is_instructional_text(field_text):
            return field_text

    # labelKey as fallback (construct the key and resolve)
    label_key = field.get("labelKey", "")
    if label_key:
        label_text = _resolve_template_label(
            f"template_label_{label_key}", defined_names
        )
        if not _is_instructional_text(label_text):
            return label_text

    # Last resort: XBRL label
    xbrl_label = field.get("label", "")
    if xbrl_label:
        return xbrl_label

    return "Unnamed field"


def _order_fields(fields: list[dict], section: dict) -> list[dict]:
    """
    Reorder fields for table sections in column-major order.

    The order proceeds from the leftmost column, row by row from the
    top, then moves to the next column.  Non-table fields (those
    without a columnLabel) are split into two groups based on their
    position relative to the table fields in the original ordering:

    * **Pre-table** — fields that appear *before* the first table
      field in the original order (e.g. yellow-cell triggers, unit
      selectors).  These are placed **before** the table data.
    * **Post-table** — fields that appear *after* the last table
      field in the original order (e.g. summary / total rows).
      These are placed **after** the table data.
    """
    col_labels = section.get("columnLabels") or {}
    row_labels = section.get("rowLabels") or {}

    if not col_labels:
        return fields  # no table structure — keep original order

    # Build reverse lookups: key → column letter, key → row number
    col_key_to_letter: dict[str, str] = {}
    for col_letter, cl in col_labels.items():
        col_key_to_letter[cl.get("key", "")] = col_letter

    row_key_to_num: dict[str, int] = {}
    for row_num_str, rl in row_labels.items():
        row_key_to_num[rl.get("key", "")] = int(row_num_str)

    # Column letter sort order (alphabetical works for single letters)
    col_order = sorted(col_labels.keys())  # e.g. ['C', 'D', 'G', 'J', 'K']
    col_rank = {letter: i for i, letter in enumerate(col_order)}

    # Identify table field positions in the original order so we can
    # split non-table fields into pre- and post-table groups.
    first_table_idx: int | None = None
    last_table_idx: int | None = None

    # Separate non-table fields (no columnLabel) from table fields
    non_table_raw: list[tuple[int, dict]] = []
    table: list[tuple[int, int, int, dict]] = []  # (col_rank, row_num, orig_idx, field)

    for orig_idx, field in enumerate(fields):
        cl = field.get("columnLabel")
        if not cl:
            non_table_raw.append((orig_idx, field))
        else:
            ck = cl.get("key", "")
            rk = (
                field.get("rowLabel", {}).get("key", "")
                if field.get("rowLabel")
                else ""
            )
            c_rank = col_rank.get(col_key_to_letter.get(ck, ""), 999)
            r_num = row_key_to_num.get(rk, 9999)
            table.append((c_rank, r_num, orig_idx, field))

            if first_table_idx is None or orig_idx < first_table_idx:
                first_table_idx = orig_idx
            if last_table_idx is None or orig_idx > last_table_idx:
                last_table_idx = orig_idx

    # Sort table fields: primary by column rank, secondary by row number
    table.sort(key=lambda t: (t[0], t[1]))

    # Split non-table fields into pre-table and post-table based on
    # their position relative to the table fields in the original order.
    pre_table: list[tuple[int, dict]] = []
    post_table: list[tuple[int, dict]] = []

    for orig_idx, field in non_table_raw:
        if last_table_idx is not None and orig_idx > last_table_idx:
            post_table.append((orig_idx, field))
        else:
            pre_table.append((orig_idx, field))

    # Pre-table (original order) → table (column-major) → post-table (original order)
    ordered = (
        [f for _, f in pre_table]
        + [f for _, _, _, f in table]
        + [f for _, f in post_table]
    )
    return ordered


def _dedup_expandable_rows(fields: list[dict], section: dict) -> list[dict]:
    """Collapse expandable-table row instances to one per unique concept.

    Sections with ``hasAdditionalRowsWarning`` contain expandable tables
    where the sample data may have multiple filled rows.  The parser emits
    one field per filled row (with different typed/explicit dimension
    members).  For survey purposes we only need **one template question**
    per unique XBRL concept in the table, so we keep the first instance
    and discard subsequent duplicates.

    Additionally, table fields that share the same ``columnLabel.key``
    (e.g. WeightOfMaterialUsed and VolumeOfMaterialUsed both mapping to
    "mass_volume") are collapsed into one question — the unit dropdown
    determines which interpretation applies at runtime.

    Non-XBRL fields (yellow cells, unit selections) are never deduplicated.
    """
    if not section.get("hasAdditionalRowsWarning"):
        return fields

    seen_qnames: set[str] = set()
    seen_col_keys: set[str] = set()
    result: list[dict] = []
    for field in fields:
        qname = field.get("qname")
        if not qname:
            # Non-XBRL fields always pass through
            result.append(field)
            continue
        if qname in seen_qnames:
            continue  # duplicate row instance — skip
        # Also collapse fields sharing the same column label key
        # (e.g. Weight + Volume in the same "mass_volume" column)
        col_label = field.get("columnLabel")
        if col_label:
            col_key = col_label.get("key", "")
            if col_key and col_key in seen_col_keys:
                continue  # alternate-unit interpretation — skip
            if col_key:
                seen_col_keys.add(col_key)
        seen_qnames.add(qname)
        result.append(field)
    return result


def _inject_synthetic_columns(fields: list[dict], section: dict) -> list[dict]:
    """Inject synthetic fields for orphaned column/index labels.

    Some expandable-table sections have column or index labels (e.g.
    "Pollutant", "Type of waste") for which the parser emitted no data
    field because the column holds a dimension axis value rather than an
    XBRL data point.  This function inserts a synthetic field for each
    such label so that the survey contains a question for it.

    Synthetic fields are defined in ``_SYNTHETIC_COLUMNS`` and inserted
    at the front of the table fields (before the first field with a
    ``columnLabel``).
    """
    section_id = section.get("sectionId", "")
    col_labels = section.get("columnLabels") or {}
    idx_labels = section.get("indexLabels") or {}

    if not col_labels and not idx_labels:
        return fields

    # Which column keys already have a matching field?
    field_col_keys: set[str] = set()
    for f in fields:
        cl = f.get("columnLabel")
        if cl:
            field_col_keys.add(cl.get("key", ""))

    # Merge all label sources (column + index) for synthetic lookup
    all_labels: dict[str, dict] = {}
    for col_letter in sorted(col_labels.keys()):
        all_labels[col_letter] = col_labels[col_letter]
    for col_letter in sorted(idx_labels.keys()):
        all_labels[f"idx_{col_letter}"] = idx_labels[col_letter]

    # Find orphaned labels that have a synthetic definition
    synthetics_to_add: list[dict] = []
    for _label_key, label_desc in all_labels.items():
        col_key = label_desc.get("key", "")
        if col_key in field_col_keys:
            continue  # already covered by an existing field

        synth_def = _SYNTHETIC_COLUMNS.get((section_id, col_key))
        if not synth_def:
            continue

        # Build a synthetic field dict with a columnLabel so it's
        # recognised as a table field by the ordering/dynamic logic.
        synth_field = dict(synth_def)
        synth_field["columnLabel"] = label_desc  # attach the label metadata
        synth_field.setdefault("rowLabel", None)
        synth_field.setdefault("value", None)
        synthetics_to_add.append(synth_field)

    if not synthetics_to_add:
        return fields

    # Insert synthetics before the first table field
    first_table_idx: int | None = None
    for i, f in enumerate(fields):
        if f.get("columnLabel"):
            first_table_idx = i
            break

    if first_table_idx is not None:
        return fields[:first_table_idx] + synthetics_to_add + fields[first_table_idx:]
    # No table fields at all — just append
    return fields + synthetics_to_add


def _is_table_unit_field(field: dict, section: dict) -> bool:
    """Return True if *field* is a consolidated unit-selection dropdown
    that belongs to the expandable table (not a section-wide unit).

    A consolidated unit field (``unitFieldIds`` present) is table-local
    only when its ``labelKey`` matches one of the section's column label
    keys.  Otherwise it's a section-wide setting (e.g. the "Please
    select the unit…" dropdown in b4 that spans the full width above
    the table header row).
    """
    if field.get("source") != "unitSelection":
        return False
    if not field.get("unitFieldIds"):
        return False

    # Check whether the unit's labelKey matches a column label key
    label_key = field.get("labelKey", "")
    col_labels = section.get("columnLabels") or {}
    for cl in col_labels.values():
        if cl.get("key", "") == label_key:
            return True

    return False


def _build_help_text(field: dict) -> str:
    """
    Compose help_text from the field's taxonomy label (which may differ from
    the template label used as the question name) and any unit information.
    """
    hints: list[str] = []

    # If the XBRL taxonomy label differs from the template label, include it
    xbrl_label = field.get("label", "")
    template_en = ""
    if field.get("templateLabels"):
        template_en = field["templateLabels"].get("en", "")
    elif field.get("rowLabel"):
        template_en = field["rowLabel"].get("templateLabels", {}).get("en", "")

    if xbrl_label and template_en and xbrl_label != template_en:
        hints.append(f"XBRL: {xbrl_label}")

    # Unit information from aspects
    aspects = field.get("aspects", {})
    units = aspects.get("units", "")
    if units:
        hints.append(f"Unit: {units}")

    decimals = aspects.get("decimals", "")
    if decimals:
        hints.append(f"Decimals: {decimals}")

    return "; ".join(hints)


def _map_answer_type(field: dict) -> str:
    """Map the report-facts inputType to a survey answer_type.

    Fields marked ``isComputed`` in the report-facts JSON are auto-
    calculated by the Excel template (their value cells have theme
    tint ≈ 0.80) and should not require user input in the survey.
    """
    if field.get("isComputed"):
        return "COMPUTED"
    input_type = field.get("inputType", "text")
    return _ANSWER_TYPE_MAP.get(input_type, "OPEN")


def _build_answer_options(
    field: dict, enum_lists: dict[str, list[str]]
) -> list[dict[str, str]] | None:
    """
    Build the answer_options list for choice-type questions.

    Sources (in priority order):
    1. domainMembers — XBRL taxonomy enumeration members
    2. options — reference to enum_* lists
    3. boolean inputType — Yes / No options
    """
    # 1) domainMembers (XBRL taxonomy members)
    if field.get("domainMembers"):
        options = []
        for dm in field["domainMembers"]:
            label = dm.get("label", "")
            # Strip " [member]" suffix from display label
            display = re.sub(r"\s*\[member\]\s*$", "", label)
            value = dm.get("qname", display)
            options.append({"value": value, "label": display})
        return options

    # 2) options — either enum_* list references or direct string values
    #    (unit selection fields use plain string lists like
    #     ["kilograms (kg)", "metric tonnes (t)"])
    if field.get("options"):
        opt_refs = field["options"]
        if isinstance(opt_refs, str):
            opt_refs = [opt_refs]

        if isinstance(opt_refs, list):
            # Check if the list contains direct values (not enum list keys)
            # by testing whether any element resolves in enum_lists.
            resolved_via_enum = False
            seen: set[str] = set()
            options: list[dict[str, str]] = []
            for ref_key in opt_refs:
                if isinstance(ref_key, str) and ref_key in enum_lists:
                    resolved_via_enum = True
                    for val in enum_lists[ref_key]:
                        if val not in seen:
                            seen.add(val)
                            options.append({"value": val, "label": val})
            if resolved_via_enum and options:
                return options

            # Not enum list references — treat as direct string values
            if not resolved_via_enum:
                direct_options = []
                for val in opt_refs:
                    if isinstance(val, str) and val not in seen:
                        seen.add(val)
                        direct_options.append({"value": val, "label": val})
                if direct_options:
                    return direct_options

    # 3) boolean → Yes/No radio
    if field.get("inputType") == "boolean":
        return [
            {"value": "Yes", "label": "Yes"},
            {"value": "No", "label": "No"},
        ]

    return None


def _build_condition(
    field: dict, section: dict, id_lookup: dict[str, str]
) -> dict[str, Any]:
    """
    Translate the validation condition from report-facts format into
    survey_data condition/condition_criteria fields.

    Returns a dict with optional keys: condition, condition_criteria,
    nesting_level.
    """
    result: dict[str, Any] = {}
    validation = field.get("validation") or {}
    condition_expr = validation.get("condition", "")
    condition_criteria = validation.get("conditionCriteria")

    if not condition_expr:
        return result

    # Bail out early if there are no condition criteria — a condition
    # without criteria is not evaluable by the survey frontend, so we
    # omit both to keep the invariant that they always appear as a pair.
    if not condition_criteria:
        return result

    # The condition expression uses {FieldName} placeholders.
    # We need to replace those with the survey question IDs.
    def _replace_ref(match: re.Match) -> str:
        ref_name = match.group(1)
        # Look up the survey question ID for this field reference
        qid = id_lookup.get(ref_name, "")
        return qid if qid else ref_name

    survey_condition = re.sub(r"\{([^}]+)\}", _replace_ref, condition_expr)

    # For simple single-field conditions, strip the enclosing braces result
    # already handled by regex above. For compound conditions, translate
    # the & to & (already correct syntax).
    result["condition"] = survey_condition

    # Build condition_criteria
    if condition_criteria and isinstance(condition_criteria, dict):
        # Multi-field criteria → translate field names to survey IDs
        translated: dict[str, str] = {}
        for ref_name, criteria_val in condition_criteria.items():
            qid = id_lookup.get(ref_name, ref_name)
            translated[qid] = criteria_val
        result["condition_criteria"] = translated
    elif condition_criteria and isinstance(condition_criteria, str):
        result["condition_criteria"] = condition_criteria

    result["nesting_level"] = 1
    return result


def _build_id_lookup(
    section: dict, fields_with_ids: list[tuple[dict, str]]
) -> dict[str, str]:
    """
    Build a mapping from field reference names (as used in validation
    conditions) to their generated survey question IDs.

    Field references can be:
    - XBRL qname local part (e.g. "BasisForPreparation")
    - template_label key (e.g. "template_label_has_the_undertaking_...")
    - yellow cell labelKey
    """
    lookup: dict[str, str] = {}
    for field, qid in fields_with_ids:
        # XBRL qname → local part (after the colon)
        qname = field.get("qname", "")
        if qname and ":" in qname:
            local = qname.split(":", 1)[1]
            lookup[local] = qid
        elif qname:
            lookup[qname] = qid

        # template_label_* key (used by yellow cell conditions)
        label_key = field.get("labelKey", "")
        if label_key:
            lookup[f"template_label_{label_key}"] = qid
            lookup[label_key] = qid

        # fieldId for yellow cells
        field_id = field.get("fieldId", "")
        if field_id:
            lookup[field_id] = qid

    return lookup


# ---------------------------------------------------------------------------
# Core conversion
# ---------------------------------------------------------------------------


def convert_field_to_question(
    field: dict,
    section: dict,
    enum_lists: dict[str, list[str]],
    question_id: str,
    id_lookup: dict[str, str],
    dynamic_config: dict[str, Any] | None = None,
    defined_names: dict[str, dict] | None = None,
) -> dict[str, Any]:
    """Convert a single report-facts field into a survey question object."""

    is_dynamic = dynamic_config is not None
    qid = f"{question_id}__DQ0" if is_dynamic else question_id

    question: dict[str, Any] = {
        "id": qid,
        "name": _build_question_name(field, section, defined_names or {}),
        "value": "",
        "answer_type": _map_answer_type(field),
        "help_text": _build_help_text(field),
        "related_ids": [],
    }

    # Answer options
    options = _build_answer_options(field, enum_lists)
    if options is not None:
        question["answer_options"] = options

    # Required — computed fields are never required (auto-calculated)
    is_required = field.get("isRequired", False)
    validation = field.get("validation") or {}
    if (is_required or validation.get("required")) and not field.get("isComputed"):
        question["required"] = True

    # Condition
    cond = _build_condition(field, section, id_lookup)
    if cond:
        question.update(cond)

    # Dynamic group fields
    if is_dynamic and dynamic_config:
        group_id = f"{SURVEY_PREFIX}_{section['sectionId']}"
        question["dynamic_group_id"] = group_id
        question["dynamic_header"] = dynamic_config.get("dynamic_header", "Item")
        question["dynamic_max_items"] = dynamic_config.get("dynamic_max_items")

        # Tag expression: use the configured field's ID as the tag
        tag_qname = dynamic_config.get("tag_field_qname")
        tag_field_key = dynamic_config.get("tag_field_key")
        if tag_qname:
            tag_id = id_lookup.get(
                tag_qname.split(":", 1)[1] if ":" in tag_qname else tag_qname, ""
            )
            if tag_id:
                question["dynamic_tag_expression"] = f"{{{tag_id}__DQ0}}"
        elif tag_field_key:
            tag_id = id_lookup.get(tag_field_key, "")
            if tag_id:
                question["dynamic_tag_expression"] = f"{{{tag_id}__DQ0}}"

    # Preserve XBRL metadata as related context (non-standard but useful)
    # Store source info for traceability
    meta: dict[str, Any] = {}
    if field.get("qname"):
        meta["xbrl_qname"] = field["qname"]
    if field.get("fieldId"):
        meta["source_field_id"] = field["fieldId"]
    if field.get("dataType"):
        meta["data_type"] = field["dataType"]
    if field.get("source") == "yellowCell":
        meta["source"] = "yellowCell"
        meta["is_reportable"] = False
    if field.get("aspects", {}).get("units"):
        meta["units"] = field["aspects"]["units"]
    if meta:
        question["_meta"] = meta

    return question


def convert_section(
    section: dict,
    enum_lists: dict[str, list[str]],
    global_id_lookup: dict[str, str],
    defined_names: dict[str, dict] | None = None,
) -> list[dict[str, Any]]:
    """
    Convert a single report-facts section into a flat list of survey
    question objects.
    """
    questions: list[dict[str, Any]] = []
    fields = section.get("fields", [])
    section_id = section["sectionId"]

    # Exclude fields that should not appear in survey output
    fields = [f for f in fields if not _should_exclude_field(section_id, f)]

    # Reorder fields for table sections (column-major order)
    fields = _order_fields(fields, section)

    # Collapse expandable-table row instances to one per concept
    fields = _dedup_expandable_rows(fields, section)

    # Inject synthetic fields for orphaned column labels
    fields = _inject_synthetic_columns(fields, section)

    # Check if this is a dynamic (repeatable) section
    dynamic_config = _DYNAMIC_SECTIONS.get(section_id)
    table_fields_only = (
        dynamic_config.get("table_fields_only", False) if dynamic_config else False
    )

    # First pass: generate IDs for all fields so we can build the lookup
    fields_with_ids: list[tuple[dict, str]] = []
    for idx, field in enumerate(fields):
        qid = _make_question_id(section_id, field, field_index=idx)
        fields_with_ids.append((field, qid))

    # Build section-local ID lookup (merged with global)
    local_lookup = _build_id_lookup(section, fields_with_ids)
    merged_lookup = {**global_id_lookup, **local_lookup}

    # Second pass: convert fields to questions
    for field, qid in fields_with_ids:
        # When table_fields_only is set, apply dynamic config only to
        # table fields (those with a columnLabel) and their associated
        # consolidated unit-selection dropdown (identified by unitFieldIds).
        effective_dynamic = dynamic_config
        if table_fields_only and dynamic_config:
            is_table_field = bool(field.get("columnLabel"))
            is_table_unit = _is_table_unit_field(field, section)
            if not (is_table_field or is_table_unit):
                effective_dynamic = None

        question = convert_field_to_question(
            field=field,
            section=section,
            enum_lists=enum_lists,
            question_id=qid,
            id_lookup=merged_lookup,
            dynamic_config=effective_dynamic,
            defined_names=defined_names,
        )
        questions.append(question)

    # Disambiguate questions that ended up with the same name but are
    # different fields (e.g. WeightOfMaterialUsed and VolumeOfMaterialUsed
    # both get column label "Mass / Volume").  Append the XBRL taxonomy
    # label in parentheses for any name that appears more than once.
    from collections import Counter

    name_counts = Counter(q["name"] for q in questions)
    for q in questions:
        if name_counts[q["name"]] > 1:
            xbrl_label = q.get("_meta", {}).get("xbrl_qname", "")
            # Use the original field's XBRL label if available
            if xbrl_label:
                # Find the original field to get its taxonomy label
                concept = xbrl_label.split(":", 1)[-1]
                # Humanise CamelCase: "WeightOfMaterialUsed" → "Weight of material used"
                human = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", concept).lower()
                human = human[0].upper() + human[1:]
                q["name"] = f"{q['name']} ({human})"

    return questions


def build_global_id_lookup(
    report: dict,
) -> dict[str, str]:
    """
    Build a global lookup mapping field reference names to generated
    question IDs, across all sections. This is needed because validation
    conditions can reference fields in other sections.
    """
    lookup: dict[str, str] = {}
    for section in report.get("sections", []):
        section_id = section.get("sectionId", "")
        raw_fields = section.get("fields", [])
        filtered_fields = [
            f for f in raw_fields if not _should_exclude_field(section_id, f)
        ]
        ordered_fields = _order_fields(filtered_fields, section)
        ordered_fields = _dedup_expandable_rows(ordered_fields, section)
        ordered_fields = _inject_synthetic_columns(ordered_fields, section)
        for idx, field in enumerate(ordered_fields):
            qid = _make_question_id(section["sectionId"], field, field_index=idx)

            qname = field.get("qname", "")
            if qname and ":" in qname:
                local = qname.split(":", 1)[1]
                lookup[local] = qid
            elif qname:
                lookup[qname] = qid

            label_key = field.get("labelKey", "")
            if label_key:
                lookup[f"template_label_{label_key}"] = qid
                lookup[label_key] = qid

            field_id = field.get("fieldId", "")
            if field_id:
                lookup[field_id] = qid

    return lookup


def _build_section_header_lookup(
    section_labels: dict[str, str],
    sections: list[dict],
) -> dict[str, str]:
    """Build a mapping from ``sectionId`` → clean section header text.

    Uses ``sectionLabels`` from report-facts.json, which is populated by
    reading the formula of each section's header cell and extracting the
    first ``template_label_*`` variable.  For example, the cell formula
    ``=template_label_b2_cooperative_specific_disclosures & " " & …``
    yields the key ``b2_cooperative_specific_disclosures`` whose resolved
    value is ``"B2 - Cooperative specific disclosures"``.

    The lookup key is the section's ``templateName`` with the leading
    ``template_`` prefix stripped (which equals the key in ``sectionLabels``).
    Falls back to a title derived from the ``sectionId`` when no match
    is found.
    """
    result: dict[str, str] = {}
    for section in sections:
        sid = section["sectionId"]
        template_name = section.get("templateName", "")
        # Strip leading "template_" to get the sectionLabels key
        key = (
            template_name[len("template_") :]
            if template_name.startswith("template_")
            else template_name
        )
        header = section_labels.get(key)
        if header:
            result[sid] = header
        else:
            # Fallback: derive from sectionId
            result[sid] = sid.replace("_", " ").strip().title()
    return result


def build_survey_wrapper(
    section: dict,
    questions: list[dict],
    section_headers: dict[str, str],
    *,
    sort_index: int,
) -> dict[str, Any]:
    """
    Wrap a section's questions in the survey object envelope.
    """
    section_id = section["sectionId"]

    # Use the section header resolved from templateLabels as the title.
    # Fall back to a title derived from the sectionId.
    title = section_headers.get(section_id) or (
        section_id.replace("_", " ").strip().title()
    )

    return {
        "title": title,
        "name": _ABBREVIATIONS[section_id],
        "sort_index": sort_index,
        "survey_group_name": section.get("sheet", ""),
        "language": "en",
        "survey_data_proxy": questions,
    }


# ---------------------------------------------------------------------------
# Sheet grouping
# ---------------------------------------------------------------------------

# Map sheet names to output filenames
_SHEET_FILE_MAP: dict[str, str] = {
    "General Information": "general_information",
    "Environmental Disclosures": "environmental_disclosures",
    "Social Disclosures": "social_disclosures",
    "Governance Disclosures": "governance_disclosures",
}


def group_sections_by_sheet(
    report: dict,
) -> dict[str, list[dict]]:
    """Group sections by their Excel sheet."""
    groups: dict[str, list[dict]] = {}
    for section in report.get("sections", []):
        sheet = section.get("sheet", "Unknown")
        if sheet not in groups:
            groups[sheet] = []
        groups[sheet].append(section)
    return groups


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    input_path = sys.argv[1] if len(sys.argv) > 1 else "output/report-facts.json"
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output/surveys"

    input_path = Path(input_path)
    output_dir = Path(output_dir)

    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Reading {input_path} ...")
    with open(input_path, "r", encoding="utf-8") as f:
        report = json.load(f)

    enum_lists = report.get("enumLists", {})
    defined_names = report.get("definedNames", {})

    # Build section header lookup from sectionLabels (formula-derived, exact)
    section_labels = report.get("sectionLabels", {})
    section_headers = _build_section_header_lookup(
        section_labels, report.get("sections", [])
    )

    # Build global cross-section ID lookup for condition references
    global_id_lookup = build_global_id_lookup(report)

    # Create one subdirectory per sheet for per-section files
    sheet_groups = group_sections_by_sheet(report)

    for sheet_name in sheet_groups:
        sheet_dir_name = _SHEET_FILE_MAP.get(
            sheet_name, sheet_name.lower().replace(" ", "_")
        )
        (output_dir / sheet_dir_name).mkdir(parents=True, exist_ok=True)

    total_questions = 0
    section_files_written = 0

    # Write one JSON file per section, placed in its sheet directory
    all_combined: list[dict[str, Any]] = []

    for sort_index, section in enumerate(report.get("sections", [])):
        section_id = section.get("sectionId", "unknown")
        sheet_name = section.get("sheet", "Unknown")
        sheet_dir_name = _SHEET_FILE_MAP.get(
            sheet_name, sheet_name.lower().replace(" ", "_")
        )

        # Build the question list for this section
        section_questions = convert_section(
            section, enum_lists, global_id_lookup, defined_names
        )
        survey_obj = build_survey_wrapper(
            section,
            section_questions,
            section_headers,
            sort_index=sort_index,
        )

        # Write per-section file into the sheet directory
        section_path = output_dir / sheet_dir_name / f"{section_id}.json"
        with open(section_path, "w", encoding="utf-8") as f:
            json.dump(survey_obj, f, indent=2, ensure_ascii=False)

        q_count = len(section_questions)
        total_questions += q_count
        section_files_written += 1
        print(f"  {section_path}: {q_count} questions")

        # Accumulate for combined file (list of survey objects)
        all_combined.append(survey_obj)

    # Write combined file as a list of survey objects
    combined_path = output_dir / "survey_data_all.json"
    with open(combined_path, "w", encoding="utf-8") as f:
        json.dump(all_combined, f, indent=2, ensure_ascii=False)
    total_q = sum(len(s["survey_data_proxy"]) for s in all_combined)
    print(
        f"  {combined_path}: {len(all_combined)} sections, {total_q} questions (combined)"
    )

    # Write a summary / index file
    total_q_all = sum(len(s["survey_data_proxy"]) for s in all_combined)
    summary: dict[str, Any] = {
        "survey_name": SURVEY_PREFIX,
        "source_file": str(input_path),
        "total_questions": total_q_all,
        "sections": {},
        "sheets": {},
    }

    for section in report.get("sections", []):
        section_id = section.get("sectionId", "unknown")
        sheet = section.get("sheet", "Unknown")
        sheet_dir_name = _SHEET_FILE_MAP.get(sheet, sheet.lower().replace(" ", "_"))
        q_count = len(section.get("fields", []))
        summary["sections"][section_id] = {
            "filename": f"{sheet_dir_name}/{section_id}.json",
            "sheet": sheet,
            "questions": q_count,
            "applicability": section.get("applicability", ""),
        }

    for sheet_name in _SHEET_FILE_MAP:
        file_key = _SHEET_FILE_MAP[sheet_name]
        if sheet_name in sheet_groups:
            sections = sheet_groups[sheet_name]
            section_ids = [s.get("sectionId") for s in sections]
            summary["sheets"][file_key] = {
                "sheet": sheet_name,
                "section_count": len(sections),
                "section_ids": section_ids,
            }

    summary_path = output_dir / "survey_index.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)

    print(
        f"\nDone. {section_files_written} section files + combined + index "
        f"written to {output_dir}/"
    )
    print(f"Total questions: {len(all_combined)}")


if __name__ == "__main__":
    main()

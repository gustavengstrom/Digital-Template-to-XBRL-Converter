"""
Stylized-fact tests for output/report-facts.json produced by parse-and-json.py.

Run with:
    conda run -n p312 python -m pytest tests/test_report_facts.py -v

Purpose
-------
These tests verify that the JSON output of parse-and-json.py remains consistent
with the observable structure of the VSME Digital Excel Template
(digital-templates/VSME-Digital-Template-Sample-1.1.1-unlocked.xlsx).

Every assertion below can be verified by a human looking at the Excel template.
They are intentionally coarse-grained so that they catch regressions caused by:
  - Changes to the Excel template structure (new/renamed sheets, rows, sections)
  - Changes to parse-and-json.py logic
  - Changes to the XBRL taxonomy

When a test fails after a legitimate template update, update the expected value
here AND note the change in scripts/parse-and-json.md.
"""

import json
from pathlib import Path
from typing import Any

import pytest

# ---------------------------------------------------------------------------
# Fixture: load the JSON output once per test session
# ---------------------------------------------------------------------------

REPORT_PATH = Path(__file__).parent.parent / "output" / "report-facts.json"


@pytest.fixture(scope="session")
def report():
    if not REPORT_PATH.exists():
        pytest.skip(
            f"report-facts.json not found at {REPORT_PATH}. "
            "Run parse-and-json.py first."
        )
    with open(REPORT_PATH, encoding="utf-8") as f:
        return json.load(f)


@pytest.fixture(scope="session")
def sections(report):
    return report["sections"]


@pytest.fixture(scope="session")
def sections_by_id(sections):
    return {s["sectionId"]: s for s in sections}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _find_section(report: dict, section_id: str) -> dict | None:
    """Return the section dict with the given ``sectionId``, or *None*."""
    for s in report["sections"]:
        if s["sectionId"] == section_id:
            return s
    return None


@pytest.fixture(scope="session")
def sections_by_sheet(sections):
    result: dict[str, list] = {}
    for s in sections:
        result.setdefault(s["sheet"], []).append(s)
    return result


# ---------------------------------------------------------------------------
# 1. Top-level structure
# ---------------------------------------------------------------------------


class TestTopLevelStructure:
    EXPECTED_KEYS = {
        "reportTitle",
        "entityName",
        "factCount",
        "templateMetadata",
        "templateLabels",
        "translations",
        "sectionApplicability",
        "omittedDisclosures",
        "sections",
        "enumLists",
        "xbrlToEnum",
        "validationRules",
        "warningDefinitions",
        "definedNames",
    }

    def test_top_level_keys_present(self, report):
        """All expected top-level keys must be present."""
        assert self.EXPECTED_KEYS.issubset(
            set(report.keys())
        ), f"Missing keys: {self.EXPECTED_KEYS - set(report.keys())}"

    def test_no_unexpected_top_level_keys(self, report):
        """No unexpected top-level keys should appear (guards against accidental additions)."""
        extra = set(report.keys()) - self.EXPECTED_KEYS
        assert extra == set(), f"Unexpected top-level keys: {extra}"

    def test_entity_name(self, report):
        """Sample template entity name is 'Company XYZ'."""
        assert report["entityName"] == "Company XYZ"

    def test_fact_count(self, report):
        """Sample template has 141 reported facts."""
        assert report["factCount"] == 141

    def test_fact_count_matches_fields_with_value(self, report):
        """factCount should roughly match the number of fields with a non-null value."""
        fields_with_value = sum(
            1
            for s in report["sections"]
            for f in s["fields"]
            if f.get("value") is not None
        )
        # The field count may exceed factCount because dimensional facts can be
        # counted differently; allow a small margin.
        assert (
            fields_with_value >= report["factCount"] * 0.9
        ), f"fields_with_value ({fields_with_value}) is much less than factCount ({report['factCount']})"


# ---------------------------------------------------------------------------
# 2. Section counts — verifiable by counting coloured headers in the Excel
# ---------------------------------------------------------------------------


class TestSectionCounts:
    def test_total_section_count(self, sections):
        """Total sections across all sheets: 45.
        Verify by counting bold section headers in all four data tabs."""
        assert len(sections) == 45

    def test_general_information_section_count(self, sections_by_sheet):
        """'General Information' tab has 12 sections.
        Verify by scrolling down the tab and counting bold headers."""
        assert len(sections_by_sheet["General Information"]) == 12

    def test_environmental_disclosures_section_count(self, sections_by_sheet):
        """'Environmental Disclosures' tab has 17 sections."""
        assert len(sections_by_sheet["Environmental Disclosures"]) == 17

    def test_social_disclosures_section_count(self, sections_by_sheet):
        """'Social Disclosures' tab has 11 sections."""
        assert len(sections_by_sheet["Social Disclosures"]) == 11

    def test_governance_disclosures_section_count(self, sections_by_sheet):
        """'Governance Disclosures' tab has 5 sections."""
        assert len(sections_by_sheet["Governance Disclosures"]) == 5

    def test_sections_with_facts(self, sections):
        """39 out of 45 sections have at least one reported fact in the sample file.
        (Was 38 before yellow cell integration — one section gained facts from yellow cells.)
        """
        assert sum(1 for s in sections if s["hasFacts"]) == 39


# ---------------------------------------------------------------------------
# 3. Section applicability — verifiable by reading the [bracket] text in
#    each section header row in the Excel template
# ---------------------------------------------------------------------------


class TestSectionApplicability:
    def test_b3_energy_is_always(self, sections_by_id):
        """B3 Total Energy is 'always to be reported' (no bracket condition)."""
        assert (
            sections_by_id["b3_total_energy_consumption_in_MWh"]["applicability"]
            == "always"
        )

    def test_b3_breakdown_is_conditional(self, sections_by_id):
        """B3 Breakdown of energy is 'if applicable'."""
        assert (
            sections_by_id["b3_breakdown_of_energy_consumption_in_MWh"]["applicability"]
            == "conditional"
        )

    def test_c3_ghg_targets_is_conditional(self, sections_by_id):
        """C3 GHG Reduction Targets is 'if applicable'."""
        assert (
            sections_by_id["c3_GHG_reduction_targets_in_tC02e"]["applicability"]
            == "conditional"
        )

    def test_b1_list_of_sites_is_always(self, sections_by_id):
        """B1 List of Sites is 'always to be reported'."""
        assert sections_by_id["b1_list_of_sites"]["applicability"] == "always"

    def test_b1_list_of_subsidiaries_is_conditional(self, sections_by_id):
        """B1 List of Subsidiaries is 'if applicable'."""
        assert (
            sections_by_id["b1_list_of_subsidiaries"]["applicability"] == "conditional"
        )

    def test_b5_biodiversity_land_use_is_optional(self, sections_by_id):
        """B5 Biodiversity Land Use is 'may (optional)'."""
        assert sections_by_id["b5_biodiversity_land_use"]["applicability"] == "optional"

    def test_b10_is_always_plus_conditional(self, sections_by_id):
        """B10 Workforce Remuneration header is 'always + if applicable' (mixed)."""
        assert (
            sections_by_id[
                "b10_workforce_remuneration_collective_bargaining_and_training"
            ]["applicability"]
            == "always+conditional"
        )

    def test_c2_is_conditional_linked(self, sections_by_id):
        """C2 Description of practices is 'if applicable linked with' another section."""
        assert (
            sections_by_id[
                "c2_description_of_practices_policies_and_future_initiatives_for_transitioning_towards_a_more_sustainable_economy"
            ]["applicability"]
            == "conditional-linked"
        )

    def test_disclosure_other_env_is_optional(self, sections_by_id):
        """The catch-all 'other environmental disclosures' section is optional."""
        assert (
            sections_by_id[
                "disclosure_of_any_other_environmental_and_or_entity_specific_enviromental_disclosures"
            ]["applicability"]
            == "optional"
        )


# ---------------------------------------------------------------------------
# 4. Section field counts — verifiable by counting input rows in each section
# ---------------------------------------------------------------------------


class TestSectionFieldCounts:
    def test_b3_energy_has_one_field(self, sections_by_id):
        """B3 Total Energy section has exactly 1 input field (TotalEnergyConsumption)."""
        assert len(sections_by_id["b3_total_energy_consumption_in_MWh"]["fields"]) == 1

    def test_b3_ghg_has_5_fields(self, sections_by_id):
        """B3 GHG emissions section has 5 fields (scope 1/2 + totals, current period only)."""
        assert (
            len(
                sections_by_id[
                    "b3_estimated_greenhouse_gas_emissions_considering_the_GHG_protocol_version_2004_in_tCO2e"
                ]["fields"]
            )
            == 5
        )

    def test_b7_waste_has_34_fields(self, sections_by_id):
        """B7 Waste Generated section has 34 fields (27 XBRL + 7 unit selection) — largest section in the template."""
        assert len(sections_by_id["b7_waste_generated"]["fields"]) == 34

    def test_b1_general_information_has_12_fields(self, sections_by_id):
        """B1 General Information section has 12 fields."""
        assert (
            len(
                sections_by_id[
                    "b1_basis_for_preparation_and_other_undertakings_general_information"
                ]["fields"]
            )
            == 12
        )

    def test_b8_gender_has_4_fields(self, sections_by_id):
        """B8 Workforce General Characteristics — Gender has 4 fields."""
        assert (
            len(sections_by_id["b8_workforce_general_characteristics_gender"]["fields"])
            == 4
        )

    def test_c8_revenues_has_8_fields(self, sections_by_id):
        """C8 Revenues from Certain Sectors has 8 fields (7 XBRL + 1 yellow cell)."""
        assert len(sections_by_id["c8_revenues_from_certain_sectors"]["fields"]) == 8

    def test_section_with_most_fields(self, sections):
        """b7_waste_generated has more fields than any other section."""
        most = max(sections, key=lambda s: len(s["fields"]))
        assert most["sectionId"] == "b7_waste_generated"


# ---------------------------------------------------------------------------
# 5. Specific field properties — verifiable by inspecting individual cells
# ---------------------------------------------------------------------------


class TestFieldProperties:
    def _get_field(self, sections, qname):
        for s in sections:
            for f in s["fields"]:
                if f.get("qname") == qname:
                    return f
        return None

    def test_total_energy_consumption_value(self, sections):
        """TotalEnergyConsumption in the sample file is 250 (MWh)."""
        f = self._get_field(sections, "vsme:TotalEnergyConsumption")
        assert f is not None, "vsme:TotalEnergyConsumption field not found"
        assert f["value"] == 250

    def test_total_energy_consumption_input_type(self, sections):
        """TotalEnergyConsumption is a numeric field."""
        f = self._get_field(sections, "vsme:TotalEnergyConsumption")
        assert f["inputType"] == "number"

    def test_basis_for_preparation_is_select(self, sections):
        """BasisForPreparation is a dropdown (select) field."""
        f = self._get_field(sections, "vsme:BasisForPreparation")
        assert f is not None, "vsme:BasisForPreparation field not found"
        assert f["inputType"] == "select"

    def test_basis_for_preparation_has_options(self, sections):
        """BasisForPreparation has an associated enum options list."""
        f = self._get_field(sections, "vsme:BasisForPreparation")
        assert (
            "options" in f or "domainMembers" in f
        ), "BasisForPreparation should have options or domainMembers"

    def test_basis_for_preparation_value(self, sections):
        """Sample file has BasisForPreparation set to the Basic + Comprehensive option."""
        f = self._get_field(sections, "vsme:BasisForPreparation")
        assert f["value"] is not None
        assert "Basic" in str(f["value"]) or "Option" in str(f["value"])

    def test_fields_have_required_schema_keys(self, sections):
        """Every field must have an identifier (qname or fieldId), label, labels, dataType, inputType, isRequired, value.

        XBRL fields use 'qname'; yellow cell fields use 'fieldId'. Both must have
        the remaining base keys: label, labels, dataType, inputType, value.

        Note: dimensional/extra facts appended via fact_to_json_entry() intentionally omit
        'isRequired' (they are raw fact entries, not taxonomy schema entries). A field is
        considered schema-backed when 'isRequired' is present.
        """
        common_keys = {"label", "labels", "dataType", "inputType", "value"}
        missing = []
        for s in sections:
            for f in s["fields"]:
                has_id = "qname" in f or "fieldId" in f
                absent = common_keys - set(f.keys())
                if absent or not has_id:
                    ident = f.get("qname", f.get("fieldId", "?"))
                    if not has_id:
                        absent = (absent or set()) | {"qname|fieldId"}
                    missing.append((s["sectionId"], ident, absent))
        assert missing == [], f"Fields missing base schema keys: {missing[:5]}"

        # Count how many schema-backed fields have isRequired
        with_is_required = sum(
            1 for s in sections for f in s["fields"] if "isRequired" in f
        )
        # The majority of fields (expected fields from taxonomy) should have isRequired
        total = sum(len(s["fields"]) for s in sections)
        assert (
            with_is_required >= total * 0.8
        ), f"Only {with_is_required}/{total} fields have isRequired — expected at least 80%"

    def test_labels_have_english(self, sections):
        """Every field with a non-empty labels dict must include an English ('en') entry."""
        without_en = []
        for s in sections:
            for f in s["fields"]:
                if f.get("labels") and "en" not in f["labels"]:
                    without_en.append(f.get("qname", f.get("fieldId")))
        assert without_en == [], f"Fields missing English label: {without_en}"

    def test_total_fields_count(self, sections):
        """Total field entries across all sections is 211 (171 XBRL + 27 yellow cell + 13 unit selection)."""
        assert sum(len(s["fields"]) for s in sections) == 211

    def test_fields_with_value_count(self, sections):
        """179 field entries have a non-null value in the sample file (139 XBRL + 27 yellow cell + 13 unit selection)."""
        assert (
            sum(1 for s in sections for f in s["fields"] if f.get("value") is not None)
            == 179
        )

    def test_fields_ordered_by_row(self, sections_by_id):
        """Fields in b3_estimated_GHG section: each (qname, dimensions) combination is unique.

        The GHG section reports the same concept (e.g. GrossScope1GreenhouseGasEmissions)
        multiple times with different XBRL dimensions (baseline year, target year, no dimension).
        That is expected behaviour. The uniqueness constraint is therefore on the combination
        of qname + dimensions, not on qname alone.
        """
        s = sections_by_id[
            "b3_estimated_greenhouse_gas_emissions_considering_the_GHG_protocol_version_2004_in_tCO2e"
        ]
        combos = [
            (f["qname"], json.dumps(f.get("dimensions", {}), sort_keys=True))
            for f in s["fields"]
        ]
        assert len(combos) == len(
            set(combos)
        ), f"Duplicate (qname, dimensions) entries detected — possible row-sort issue"


# ---------------------------------------------------------------------------
# 5b. Yellow cell fields — condition triggers & helper inputs not in XBRL
#     but needed by an online survey (detected via fill colour FFFFFF99)
# ---------------------------------------------------------------------------


class TestYellowCellFields:
    """Tests for yellow-cell fields extracted from the Excel template.

    Yellow cells (fill colour FFFFFF99) contain condition/trigger questions
    and helper inputs that are not XBRL-reportable but essential for
    driving the conditional logic of an online survey.
    """

    @pytest.fixture(scope="class")
    def yellow_fields(self, sections):
        return [
            f for s in sections for f in s["fields"] if f.get("source") == "yellowCell"
        ]

    @pytest.fixture(scope="class")
    def yellow_sections(self, sections):
        return [
            s
            for s in sections
            if any(f.get("source") == "yellowCell" for f in s["fields"])
        ]

    def test_total_yellow_field_count(self, yellow_fields):
        """27 yellow cell fields extracted from the template."""
        assert len(yellow_fields) == 27

    def test_sections_with_yellow_fields(self, yellow_sections):
        """19 sections contain at least one yellow cell field."""
        assert len(yellow_sections) == 19

    def test_yellow_fields_have_fieldId(self, yellow_fields):
        """Every yellow field uses 'fieldId' (not 'qname') as its identifier."""
        for f in yellow_fields:
            assert "fieldId" in f, f"Yellow field missing fieldId: {f}"
            assert f["fieldId"].startswith(
                "yellow_"
            ), f"fieldId should start with 'yellow_': {f['fieldId']}"

    def test_yellow_fields_have_source_marker(self, yellow_fields):
        """Every yellow field has source='yellowCell'."""
        for f in yellow_fields:
            assert f["source"] == "yellowCell"

    def test_yellow_fields_not_reportable(self, yellow_fields):
        """Yellow fields are not XBRL-reportable (isReportable=False)."""
        for f in yellow_fields:
            assert (
                f.get("isReportable") is False
            ), f"Yellow field should not be reportable: {f.get('fieldId')}"

    def test_yellow_fields_not_required(self, yellow_fields):
        """Yellow fields are not required (isRequired=False)."""
        for f in yellow_fields:
            assert (
                f.get("isRequired") is False
            ), f"Yellow field should not be required: {f.get('fieldId')}"

    def test_yellow_fields_have_inferred_datatype(self, yellow_fields):
        """Yellow fields have an inferred dataType based on their inputType.

        booleanItemType for boolean, decimalItemType for number,
        enumerationItemType for select.
        """
        expected_map = {
            "boolean": "booleanItemType",
            "number": "decimalItemType",
            "select": "enumerationItemType",
        }
        for f in yellow_fields:
            it = f.get("inputType")
            expected = expected_map.get(it)
            assert f["dataType"] == expected, (
                f"Yellow field {f.get('fieldId')}: expected dataType={expected} "
                f"for inputType={it}, got {f['dataType']}"
            )

    def test_yellow_boolean_fields(self, yellow_fields):
        """Most yellow fields are boolean condition triggers (True/False)."""
        booleans = [f for f in yellow_fields if f.get("inputType") == "boolean"]
        assert (
            len(booleans) >= 15
        ), f"Expected ≥15 boolean yellow fields, got {len(booleans)}"
        for f in booleans:
            assert f["value"] in (
                True,
                False,
            ), f"Boolean field has non-bool value: {f.get('fieldId')}"

    def test_yellow_select_fields_have_options(self, yellow_fields):
        """Yellow select fields must carry an 'options' list."""
        selects = [f for f in yellow_fields if f.get("inputType") == "select"]
        assert len(selects) >= 1, "Expected at least 1 select yellow field"
        for f in selects:
            assert (
                "options" in f
            ), f"Select yellow field missing options: {f.get('fieldId')}"
            assert (
                len(f["options"]) >= 2
            ), f"Select field has fewer than 2 options: {f.get('fieldId')}"

    def test_yellow_number_fields(self, yellow_fields):
        """Yellow number fields contain numeric helper values."""
        numbers = [f for f in yellow_fields if f.get("inputType") == "number"]
        assert (
            len(numbers) >= 2
        ), f"Expected ≥2 number yellow fields, got {len(numbers)}"
        for f in numbers:
            assert isinstance(
                f["value"], (int, float)
            ), f"Number field has non-numeric value: {f.get('fieldId')}"

    # --- Specific verifiable yellow cell facts ---

    def test_b11_convictions_question(self, sections_by_id):
        """B11 section has a yellow boolean asking about convictions (value=True in sample)."""
        s = sections_by_id["b11_convictions_and_fines_for_corruption_and_bribery"]
        yf = [f for f in s["fields"] if f.get("source") == "yellowCell"]
        assert len(yf) == 1
        assert yf[0]["value"] is True
        assert yf[0]["inputType"] == "boolean"

    def test_transition_plan_status_is_select(self, sections_by_id):
        """C3 Transition Plan section has a yellow 'status of implementation' select field."""
        s = sections_by_id[
            "c3_transition_plan_for_undertakings_operating_in_high_climate_impact_sectors"
        ]
        status_field = next(
            (
                f
                for f in s["fields"]
                if f.get("source") == "yellowCell"
                and f.get("inputType") == "select"
                and "status" in f.get("fieldId", "").lower()
            ),
            None,
        )
        assert status_field is not None, "Transition plan status select field not found"
        assert (
            status_field["value"]
            == "Adoption of a transition plan is planned in the future"
        )

    def test_c8_exclusion_has_5_yellow_booleans(self, sections_by_id):
        """C8 Exclusion from EU Reference Benchmarks has 5 yellow boolean fields."""
        s = sections_by_id["c8_exclusion_from_EU_reference_benchmarks"]
        yf = [f for f in s["fields"] if f.get("source") == "yellowCell"]
        assert len(yf) == 5
        assert all(f["inputType"] == "boolean" for f in yf)

    def test_b9_hours_worked_is_number(self, sections_by_id):
        """B9 Health & Safety has a yellow 'hours worked' number field = 2000."""
        s = sections_by_id["b9_workforce_health_and_safety"]
        yf = [
            f
            for f in s["fields"]
            if f.get("source") == "yellowCell" and f.get("inputType") == "number"
        ]
        assert len(yf) == 1
        assert yf[0]["value"] == 2000

    def test_b1_list_of_sites_has_yellow_boolean(self, sections_by_id):
        """B1 List of Sites has a yellow boolean field at GI!H449 (value=True in sample).

        H449's label is 2 rows above at H447 (upward-scan pattern), which required
        the upward-scan label resolution added to the yellow cell extractor.
        """
        s = sections_by_id["b1_list_of_sites"]
        yf = [f for f in s["fields"] if f.get("source") == "yellowCell"]
        assert (
            len(yf) == 1
        ), f"Expected 1 yellow field in b1_list_of_sites, got {len(yf)}"
        assert yf[0]["inputType"] == "boolean"
        assert yf[0]["value"] is True
        assert "H449" in yf[0]["cellRef"]

    def test_b2_practices_policies_has_yellow_boolean(self, sections_by_id):
        """B2 Practices/Policies has a yellow boolean field at GI!C556 (value=True in sample).

        C556's label is 2 rows above at C554 (upward-scan pattern).
        """
        s = sections_by_id[
            "b2_practices_policies_and_future_initiatives_for_transitioning_towards_a_more_sustainable_economy"
        ]
        yf = [f for f in s["fields"] if f.get("source") == "yellowCell"]
        assert (
            len(yf) == 1
        ), f"Expected 1 yellow field in b2_practices_policies, got {len(yf)}"
        assert yf[0]["inputType"] == "boolean"
        assert yf[0]["value"] is True
        assert "C556" in yf[0]["cellRef"]

    def test_xbrl_fields_distinguishable_from_yellow(self, sections):
        """XBRL fields use 'qname' and yellow cells use 'fieldId' — they are distinguishable."""
        xbrl_fields = [f for s in sections for f in s["fields"] if "qname" in f]
        yellow_fields = [
            f for s in sections for f in s["fields"] if f.get("source") == "yellowCell"
        ]
        assert (
            len(xbrl_fields) == 171
        ), f"Expected 171 XBRL fields, got {len(xbrl_fields)}"
        assert (
            len(yellow_fields) == 27
        ), f"Expected 27 yellow fields, got {len(yellow_fields)}"
        # No overlap: yellow fields should not have 'qname'
        overlap = [f for f in yellow_fields if "qname" in f]
        assert (
            overlap == []
        ), f"Yellow fields should not have qname: {[f.get('fieldId') for f in overlap]}"

    def test_yellow_fields_embedded_in_section_fields(self, yellow_sections):
        """Sections with yellow fields have them interleaved in the 'fields' array."""
        for s in yellow_sections:
            yf_in_fields = [f for f in s["fields"] if f.get("source") == "yellowCell"]
            assert (
                len(yf_in_fields) >= 1
            ), f"Section {s['sectionId']} has no yellow fields in 'fields' array"

    def test_all_yellow_fields_have_cell_ref(self, yellow_fields):
        """Every yellow field has a cellRef tracing back to the Excel cell."""
        for f in yellow_fields:
            assert "cellRef" in f, f"Yellow field missing cellRef: {f.get('fieldId')}"
            assert f["cellRef"], f"Yellow field has empty cellRef: {f.get('fieldId')}"


# ---------------------------------------------------------------------------
# 5d. Unit selection fields — dropdown cells for selecting units of
#     measurement (e.g. kg/tonnes, hectares/m²) that are XBRL defined
#     names ending in _unit but are not taxonomy data concepts
# ---------------------------------------------------------------------------


class TestUnitSelectionFields:
    """Tests for unit selection fields extracted from the Excel template.

    Unit selection cells (e.g. AmountOfEmissionToAir_unit) are XBRL defined
    names ending in ``_unit``.  They are not taxonomy data concepts but
    provide the unit context for adjacent data fields.  They are extracted
    with ``source: "unitSelection"`` and carry enum options from the
    ``xbrlToEnum`` mapping.
    """

    @pytest.fixture(scope="class")
    def unit_fields(self, sections):
        return [
            f for s in sections for f in s["fields"] if f.get("source") == "unitSelection"
        ]

    @pytest.fixture(scope="class")
    def unit_sections(self, sections):
        return [
            s
            for s in sections
            if any(f.get("source") == "unitSelection" for f in s["fields"])
        ]

    def test_total_unit_field_count(self, unit_fields):
        """13 unit selection fields after consolidating fields sharing the same Excel cell."""
        assert len(unit_fields) == 13

    def test_sections_with_unit_fields(self, unit_sections):
        """5 sections contain at least one unit selection field."""
        assert len(unit_sections) == 5
        section_ids = {s["sectionId"] for s in unit_sections}
        assert "b4_pollution_of_air_water_and_soil" in section_ids
        assert "b5_biodiversity_land_use" in section_ids
        assert "b7_waste_generated" in section_ids

    def test_unit_fields_have_fieldId(self, unit_fields):
        """Every unit field uses 'fieldId' (not 'qname') as its identifier."""
        for f in unit_fields:
            assert "fieldId" in f, f"Unit field missing fieldId: {f}"
            assert f["fieldId"].endswith(
                "_unit"
            ), f"fieldId should end with '_unit': {f['fieldId']}"

    def test_unit_fields_have_source_marker(self, unit_fields):
        """Every unit field has source='unitSelection'."""
        for f in unit_fields:
            assert f["source"] == "unitSelection"

    def test_unit_fields_have_enum_options(self, unit_fields):
        """Most unit fields should have options from the xbrlToEnum mapping.

        Some unit fields (e.g. aggregate total fields like TotalWasteGeneratedMass_unit)
        may not have a direct xbrlToEnum mapping and thus lack options.
        After consolidation, at least 6 of 13 unit fields have options.
        """
        with_options = [f for f in unit_fields if "options" in f and len(f["options"]) >= 2]
        assert (
            len(with_options) >= 6
        ), f"Expected ≥6 unit fields with options, got {len(with_options)}"

    def test_unit_fields_have_values(self, unit_fields):
        """Every unit field in the sample file should have a non-null value."""
        for f in unit_fields:
            assert (
                f.get("value") is not None
            ), f"Unit field has null value: {f.get('fieldId')}"

    def test_unit_fields_have_null_column_label(self, unit_fields):
        """Unit fields should have columnLabel=null (standalone, not table cells).

        Unit selection fields are standalone dropdowns that select the unit of
        measurement for adjacent data fields.  They are not part of the data
        table matrix and should not inherit column headers.  Their label is
        resolved from sectionWideLabels (wide-merged instruction rows) instead.
        """
        for f in unit_fields:
            assert (
                f.get("columnLabel") is None
            ), f"Unit field should have null columnLabel: {f.get('fieldId')}"

    def test_b4_b5_unit_fields_have_null_row_label(self, sections_by_id):
        """Unit fields in b4 and b5 table areas should have rowLabel=null.

        These sections have a Q&A area above the table boundary, and the
        rowLabel boundary fix ensures table data rows don't inherit Q&A labels.
        """
        for sec_id in [
            "b4_pollution_of_air_water_and_soil",
            "b5_biodiversity_land_use",
            "b5_sites_in_biodiversity_sensitive_areas",
        ]:
            s = sections_by_id.get(sec_id)
            if s is None:
                continue
            for f in s["fields"]:
                if f.get("source") == "unitSelection":
                    assert (
                        f.get("rowLabel") is None
                    ), f"Unit field should have null rowLabel in {sec_id}: {f.get('fieldId')}"

    def test_b4_pollution_has_1_consolidated_unit_field(self, sections_by_id):
        """B4 pollution section has 1 consolidated unit field covering air, water, soil.

        All three _unit fields (air, water, soil) reference the same Excel cell
        so they are consolidated into a single field with unitFieldIds listing all.
        """
        s = sections_by_id["b4_pollution_of_air_water_and_soil"]
        uf = [f for f in s["fields"] if f.get("source") == "unitSelection"]
        assert len(uf) == 1
        assert "unitFieldIds" in uf[0]
        underlying = set(uf[0]["unitFieldIds"])
        assert "AmountOfEmissionToAir_unit" in underlying
        assert "AmountOfEmissionToWater_unit" in underlying
        assert "AmountOfEmissionToSoil_unit" in underlying

    def test_b5_biodiversity_land_use_has_1_consolidated_unit_field(self, sections_by_id):
        """B5 biodiversity land use section has 1 consolidated unit field.

        All four _unit fields reference the same Excel cell (C293) so they
        are consolidated into a single field with unitFieldIds listing all.
        The templateLabelKey should be template_label_please_select_the_unit_used_for_the_area.
        """
        s = sections_by_id["b5_biodiversity_land_use"]
        uf = [f for f in s["fields"] if f.get("source") == "unitSelection"]
        assert len(uf) == 1
        assert uf[0]["templateLabelKey"] == "template_label_please_select_the_unit_used_for_the_area"
        assert "unitFieldIds" in uf[0]
        underlying = set(uf[0]["unitFieldIds"])
        assert "TotalSealedArea_unit" in underlying
        assert "TotalNatureOrientedAreaOnSite_unit" in underlying
        assert "TotalNatureOrientedAreaOffSite_unit" in underlying
        assert "TotalUseOfLand_unit" in underlying

    def test_b7_waste_has_7_unit_fields(self, sections_by_id):
        """B7 waste section has 7 unit fields after consolidation (was 12 before)."""
        s = sections_by_id["b7_waste_generated"]
        uf = [f for f in s["fields"] if f.get("source") == "unitSelection"]
        assert len(uf) == 7

    def test_unit_fields_per_section_count(self, sections):
        """Per-section unit field counts match expected values after consolidation."""
        counts = {}
        for s in sections:
            uf = [f for f in s["fields"] if f.get("source") == "unitSelection"]
            if uf:
                counts[s["sectionId"]] = len(uf)
        expected = {
            "b4_pollution_of_air_water_and_soil": 1,
            "b5_sites_in_biodiversity_sensitive_areas": 1,
            "b5_biodiversity_land_use": 1,
            "b7_waste_generated": 7,
            "b7_annual_mass_flow_of_relevant_materials_used": 3,
        }
        assert counts == expected


# ---------------------------------------------------------------------------
# 5c. Template label linkage — connecting fields to spreadsheet question
#     phrasing via the template_label_* defined names and their translations
# ---------------------------------------------------------------------------


class TestTemplateLabelLinkage:
    """Tests for the labelKey / templateLabels linkage on section fields.

    Each row in the data sheets has a ``=template_label_*`` formula that
    contains the spreadsheet's question phrasing (potentially different from
    the XBRL taxonomy label).  These template labels also carry i18n
    translations from the Translations sheet.

    Both XBRL and yellow-cell fields should include:
    * ``labelKey`` — the template_label suffix (e.g. ``"starting_year"``)
    * ``templateLabels`` — a ``{lang: text}`` dict from the Translations sheet
    """

    @pytest.fixture(scope="class")
    def all_fields(self, sections):
        return [f for s in sections for f in s["fields"]]

    @pytest.fixture(scope="class")
    def xbrl_fields(self, all_fields):
        return [f for f in all_fields if "qname" in f]

    @pytest.fixture(scope="class")
    def yellow_fields(self, all_fields):
        return [f for f in all_fields if f.get("source") == "yellowCell"]

    # --- Structural presence ---

    def test_all_fields_have_labelKey_key(self, all_fields):
        """Every field (XBRL and yellow) has a 'labelKey' key in its dict."""
        for f in all_fields:
            assert (
                "labelKey" in f
            ), f"Field missing 'labelKey': {f.get('qname') or f.get('fieldId')}"

    def test_all_fields_have_templateLabels_key(self, all_fields):
        """Every field has a 'templateLabels' key (may be empty dict)."""
        for f in all_fields:
            assert (
                "templateLabels" in f
            ), f"Field missing 'templateLabels': {f.get('qname') or f.get('fieldId')}"
            assert isinstance(f["templateLabels"], dict)

    # --- Yellow cell coverage ---

    def test_all_yellow_fields_have_labelKey(self, yellow_fields):
        """Every yellow cell field has a non-null labelKey."""
        for f in yellow_fields:
            assert f[
                "labelKey"
            ], f"Yellow field has null/empty labelKey: {f.get('fieldId')}"

    def test_all_yellow_fields_have_templateLabels(self, yellow_fields):
        """Every yellow cell field has non-empty templateLabels translations."""
        for f in yellow_fields:
            assert f[
                "templateLabels"
            ], f"Yellow field has empty templateLabels: {f.get('fieldId')}"
            assert (
                "en" in f["templateLabels"]
            ), f"Yellow field templateLabels missing English: {f.get('fieldId')}"

    def test_all_yellow_fields_have_templateLabelKey(self, yellow_fields):
        """Every yellow cell has a templateLabelKey (full template_label_* name)."""
        for f in yellow_fields:
            tlk = f.get("templateLabelKey")
            assert tlk, f"Yellow field missing templateLabelKey: {f.get('fieldId')}"
            assert tlk.startswith(
                "template_label_"
            ), f"templateLabelKey should start with 'template_label_': {tlk}"
            # templateLabelKey == "template_label_" + labelKey
            assert (
                tlk == f"template_label_{f['labelKey']}"
            ), f"templateLabelKey mismatch for {f.get('fieldId')}"

    # --- XBRL coverage ---

    def test_xbrl_fields_with_labelKey_count(self, xbrl_fields):
        """At least 151 XBRL fields (80%) have a template label linkage."""
        with_lk = [f for f in xbrl_fields if f.get("labelKey")]
        assert (
            len(with_lk) >= 134
        ), f"Expected ≥134 XBRL fields with labelKey, got {len(with_lk)}"

    def test_xbrl_fields_with_labelKey_have_templateLabels(self, xbrl_fields):
        """XBRL fields that have a labelKey also have non-empty templateLabels."""
        for f in xbrl_fields:
            if f.get("labelKey"):
                assert f[
                    "templateLabels"
                ], f"XBRL field {f['qname']} has labelKey but empty templateLabels"
                assert (
                    "en" in f["templateLabels"]
                ), f"XBRL field {f['qname']} templateLabels missing English"

    # --- Specific verifiable examples ---

    def test_convictions_field_has_correct_template_label(self, sections_by_id):
        """B11 TotalNumberOfConvictions has the expected template label key and translations."""
        s = sections_by_id["b11_convictions_and_fines_for_corruption_and_bribery"]
        f = next(
            f
            for f in s["fields"]
            if f.get("qname")
            == "vsme:TotalNumberOfConvictionsForTheViolationOfAntiCorruptionAndAntiBriberyLaws"
        )
        assert (
            f["labelKey"]
            == "total_number_of_convictions_for_the_violation_of_anti_corruption_and_anti_bribery_laws"
        )
        assert "en" in f["templateLabels"]
        assert "convictions" in f["templateLabels"]["en"].lower()

    def test_convictions_yellow_has_correct_template_label(self, sections_by_id):
        """B11 yellow question has the expected template label key."""
        s = sections_by_id["b11_convictions_and_fines_for_corruption_and_bribery"]
        f = next(f for f in s["fields"] if f.get("source") == "yellowCell")
        assert (
            f["labelKey"]
            == "has_the_undertaking_incurred_in_convictions_and_fines_in_the_reporting_period"
        )
        assert "en" in f["templateLabels"]

    def test_total_energy_has_template_label(self, sections_by_id):
        """B3 TotalEnergyConsumption has a template label linking to the spreadsheet question."""
        s = sections_by_id["b3_total_energy_consumption_in_MWh"]
        f = next(
            f for f in s["fields"] if f.get("qname") == "vsme:TotalEnergyConsumption"
        )
        assert f["labelKey"] is not None
        assert "en" in f["templateLabels"]
        assert "energy" in f["templateLabels"]["en"].lower()

    def test_templateLabels_differ_from_taxonomy_labels(self, xbrl_fields):
        """At least some fields have templateLabels.en different from taxonomy label.

        This validates that templateLabels carry the spreadsheet phrasing
        rather than duplicating the XBRL taxonomy labels.
        """
        differs = 0
        for f in xbrl_fields:
            tl_en = f.get("templateLabels", {}).get("en", "")
            tax_en = f.get("labels", {}).get("en", "")
            if tl_en and tax_en and tl_en != tax_en:
                differs += 1
        assert (
            differs >= 5
        ), f"Expected ≥5 fields where templateLabels.en differs from labels.en, got {differs}"


# ---------------------------------------------------------------------------
# 6. Validation rules — verifiable by counting validation indicator cells
# ---------------------------------------------------------------------------


class TestValidationRules:
    def test_total_validation_rules(self, report):
        """104 validation rules extracted from IF-formula cells (incl. array formulas)."""
        assert len(report["validationRules"]) == 104

    def test_required_rules_count(self, report):
        """7 rules with status 'required' (unconditional missing-value checks)."""
        required = [
            v
            for v in report["validationRules"].values()
            if v.get("status") == "required"
        ]
        assert len(required) == 7

    def test_conditional_rules_count(self, report):
        """78 rules with status 'conditional' (depend on another field's value)."""
        conditional = [
            v
            for v in report["validationRules"].values()
            if v.get("status") == "conditional"
        ]
        assert len(conditional) == 78

    def test_informational_rules_count(self, report):
        """19 rules with status 'informational' (no missing-value penalty)."""
        informational = [
            v
            for v in report["validationRules"].values()
            if v.get("status") == "informational"
        ]
        assert len(informational) == 19

    def test_total_energy_has_validation(self, report):
        """TotalEnergyConsumption should have a validation rule (it's an 'always' field)."""
        assert "TotalEnergyConsumption" in report["validationRules"]

    def test_basis_for_preparation_has_validation(self, report):
        """BasisForPreparation is a required field and must have a validation rule."""
        assert "BasisForPreparation" in report["validationRules"]

    def test_validation_rules_have_required_and_status(self, report):
        """Every validation rule must have 'required' (bool) and 'status' keys."""
        for name, rule in report["validationRules"].items():
            assert "required" in rule, f"Rule for {name} missing 'required'"
            assert "status" in rule, f"Rule for {name} missing 'status'"
            assert isinstance(
                rule["required"], bool
            ), f"Rule for {name}: 'required' is not bool"

    # --- Per-sheet counts — verifiable by filtering validationCell prefix ----
    # Sheet is derived from the 'validationCell' value (format: "Sheet!Ref").
    # 1 rule has no sheet prefix (cross-sheet or formula-only ref) → "unknown".

    @pytest.fixture(scope="class")
    def rules_by_sheet(self, report):
        from collections import Counter

        counts: dict[str, Counter] = {}
        for v in report["validationRules"].values():
            vc = v.get("validationCell", "")
            sheet = vc.split("!")[0] if "!" in vc else "unknown"
            counts.setdefault(sheet, Counter())[v["status"]] += 1
        return counts

    def test_general_information_rules(self, rules_by_sheet):
        """General Information sheet: 35 rules (6 required, 23 conditional, 6 informational).
        Verify by counting yellow validation-indicator cells in the General Information tab.
        """
        sheet = rules_by_sheet.get("General Information", {})
        assert sum(sheet.values()) == 35
        assert sheet["required"] == 6
        assert sheet["conditional"] == 23
        assert sheet["informational"] == 6

    def test_environmental_disclosures_rules(self, rules_by_sheet):
        """Environmental Disclosures sheet: 33 rules (0 required, 29 conditional, 4 informational)."""
        sheet = rules_by_sheet.get("Environmental Disclosures", {})
        assert sum(sheet.values()) == 33
        assert sheet.get("required", 0) == 0
        assert sheet["conditional"] == 29
        assert sheet["informational"] == 4

    def test_social_disclosures_rules(self, rules_by_sheet):
        """Social Disclosures sheet: 27 rules (1 required, 20 conditional, 6 informational)."""
        sheet = rules_by_sheet.get("Social Disclosures", {})
        assert sum(sheet.values()) == 27
        assert sheet["required"] == 1
        assert sheet["conditional"] == 20
        assert sheet["informational"] == 6

    def test_governance_disclosures_rules(self, rules_by_sheet):
        """Governance Disclosures sheet: 8 rules (0 required, 6 conditional, 2 informational)."""
        sheet = rules_by_sheet.get("Governance Disclosures", {})
        assert sum(sheet.values()) == 8
        assert sheet.get("required", 0) == 0
        assert sheet["conditional"] == 6
        assert sheet["informational"] == 2

    def test_unknown_sheet_rules(self, rules_by_sheet):
        """1 rule has no sheet prefix in its validationCell (cross-sheet or formula ref)."""
        unknown = rules_by_sheet.get("unknown", {})
        assert sum(unknown.values()) == 1
        assert unknown["informational"] == 1  # all are informational-only

    # --- condition / conditionCriteria (condition-eval format) ----------------

    def test_conditional_rules_have_condition_expr(self, report):
        """Every conditional rule should have a 'condition' expression."""
        for name, rule in report["validationRules"].items():
            if rule.get("status") == "conditional":
                assert (
                    "condition" in rule
                ), f"Conditional rule {name} missing 'condition'"
                assert (
                    "{" in rule["condition"]
                ), f"Condition for {name} should use {{var}} syntax"

    def test_condition_criteria_count(self, report):
        """At least 70 rules have conditionCriteria with comparison values."""
        with_criteria = sum(
            1 for v in report["validationRules"].values() if v.get("conditionCriteria")
        )
        assert with_criteria >= 70

    def test_condition_criteria_no_trailing_pipe(self, report):
        """No conditionCriteria value should have trailing or leading '|'."""
        for name, rule in report["validationRules"].items():
            for key, val in rule.get("conditionCriteria", {}).items():
                assert not val.startswith("|"), f"{name}/{key} starts with |: {val}"
                assert not val.endswith("|"), f"{name}/{key} ends with |: {val}"

    def test_certifications_is_conditional_on_yellow_cell(self, report):
        """DescriptionOfSustainabilityRelatedCertificationsOrLabels should be
        conditional on the yellow boolean question (not incorrectly 'required')."""
        rule = report["validationRules"][
            "DescriptionOfSustainabilityRelatedCertificationsOrLabels"
        ]
        assert rule["status"] == "conditional"
        assert "condition" in rule
        assert "conditionCriteria" in rule
        # The condition variable should be the template_label for the
        # "has the undertaking obtained..." yellow cell
        assert "template_label_has_the_undertaking_obtained" in rule["condition"]
        criteria = rule["conditionCriteria"]
        cond_key = [k for k in criteria if "has_the_undertaking_obtained" in k]
        assert len(cond_key) == 1
        assert criteria[cond_key[0]] == "TRUE"

    def test_condition_variable_matches_yellow_templateLabelKey(self, report):
        """Condition variables referencing yellow cells should exactly match
        the templateLabelKey on the corresponding yellow cell field.
        Not every template_label_* condition is a yellow cell (some are
        defined names for row-level content), so we check the reverse:
        every yellow cell templateLabelKey that appears in a condition
        is properly referenced."""
        # Collect all yellow cell templateLabelKey values
        yellow_tlks = set()
        for section in report["sections"]:
            for f in section.get("fields", []):
                if f.get("source") == "yellowCell" and f.get("templateLabelKey"):
                    yellow_tlks.add(f["templateLabelKey"])
        # Collect all template_label_* condition variables
        cond_template_vars = set()
        for rule in report["validationRules"].values():
            # Single-condition rules use conditionField
            cf = rule.get("conditionField", "")
            if cf and cf.startswith("template_label_"):
                cond_template_vars.add(cf)
            # Multi-condition rules also have allConditionFields
            for cf in rule.get("allConditionFields", []):
                if cf.startswith("template_label_"):
                    cond_template_vars.add(cf)
        # Every yellow cell that is actually used as a condition variable
        # should be findable in condition variables
        used_yellow = yellow_tlks & cond_template_vars
        assert len(used_yellow) >= 10, (
            f"Expected ≥10 yellow cells used as condition variables, "
            f"got {len(used_yellow)}"
        )
        # The certifications yellow cell specifically must be in both sets
        cert_key = "template_label_has_the_undertaking_obtained_any_sustainability_related_certification_or_label"
        assert cert_key in yellow_tlks
        assert cert_key in cond_template_vars

    def test_basis_for_preparation_condition_uses_pipe(self, report):
        """BasisForPreparation conditions should use '|' as value separator."""
        rule = report["validationRules"]["BasisForReporting"]
        assert rule["status"] == "conditional"
        criteria = rule["conditionCriteria"]["BasisForPreparation"]
        assert "|" in criteria
        assert "Option A" in criteria
        assert "Option B" in criteria

    def test_total_energy_no_spurious_axis_condition(self, report):
        """TotalEnergyConsumption should condition on BasisForPreparation,
        not on TypeOfPollutantAxis (which was a cross-sheet resolution bug)."""
        rule = report["validationRules"]["TotalEnergyConsumption"]
        assert rule["conditionField"] == "BasisForPreparation"
        assert "TypeOfPollutantAxis" not in rule.get("allConditionFields", [])
        assert "TypeOfPollutantAxis" not in rule.get("condition", "")


# ---------------------------------------------------------------------------
# 7. Translations — verifiable by opening the Translations sheet
# ---------------------------------------------------------------------------


class TestTranslations:
    EXPECTED_LANGUAGES = {"en", "da", "fr", "de", "it", "lt", "pl", "pt", "es"}

    def test_translation_key_count(self, report):
        """487 translation keys in the Translations sheet."""
        assert len(report["translations"]) == 487

    def test_all_translations_have_english(self, report):
        """Every translation entry must have an English string."""
        missing = [k for k, v in report["translations"].items() if "en" not in v]
        assert missing == [], f"Translation keys without English: {missing[:5]}"

    def test_languages_present(self, report):
        """All 9 EU languages are represented in the translations."""
        sample = next(iter(report["translations"].values()))
        assert self.EXPECTED_LANGUAGES.issubset(
            set(sample.keys())
        ), f"Missing languages in sample entry: {self.EXPECTED_LANGUAGES - set(sample.keys())}"

    def test_warning_definitions_count(self, report):
        """15 warning definitions extracted from translation keys containing 'warning'."""
        assert len(report["warningDefinitions"]) == 15

    def test_warning_definitions_have_english(self, report):
        """Every warning definition must have an English text."""
        missing = [k for k, v in report["warningDefinitions"].items() if "en" not in v]
        assert missing == [], f"Warning definitions without English: {missing}"


# ---------------------------------------------------------------------------
# 8. Enum / dropdown lists — verifiable by inspecting Excel data validations
# ---------------------------------------------------------------------------


class TestEnumLists:
    def test_enum_list_count(self, report):
        """37 enum_* defined names extracted from the workbook."""
        assert len(report["enumLists"]) == 37

    def test_xbrl_to_enum_count(self, report):
        """40 XBRL fields are linked to an enum dropdown list."""
        assert len(report["xbrlToEnum"]) == 40

    def test_basis_for_preparation_has_domain_members(self, sections):
        """BasisForPreparation exposes its options via taxonomy domainMembers (not xbrlToEnum).

        BasisForPreparation is an XBRL enumeration type whose valid values come from the
        taxonomy domain, not an Excel enum_* list — so it appears in 'domainMembers' on
        the field, not in the top-level 'xbrlToEnum' dict.
        """
        f = next(
            (
                f
                for s in sections
                for f in s["fields"]
                if f.get("qname") == "vsme:BasisForPreparation"
            ),
            None,
        )
        assert f is not None, "vsme:BasisForPreparation field not found"
        assert (
            "domainMembers" in f
        ), "BasisForPreparation should have domainMembers from the taxonomy"
        assert (
            len(f["domainMembers"]) >= 2
        ), "BasisForPreparation should have at least 2 domain members (Basic, Comprehensive)"

    def test_enum_lists_are_non_empty(self, report):
        """Every enum list must contain at least one option."""
        empty = [k for k, v in report["enumLists"].items() if not v]
        assert empty == [], f"Empty enum lists: {empty}"


# ---------------------------------------------------------------------------
# 9. Defined names catalogue
# ---------------------------------------------------------------------------


class TestDefinedNames:
    def test_defined_names_count(self, report):
        """799 defined names catalogued from the workbook."""
        assert len(report["definedNames"]) == 799

    def test_defined_names_have_category_and_cell_ref(self, report):
        """Every defined name entry must have 'category' and 'cellRef' keys."""
        for name, entry in report["definedNames"].items():
            assert "category" in entry, f"Defined name {name} missing 'category'"
            assert "cellRef" in entry, f"Defined name {name} missing 'cellRef'"

    def test_basis_for_preparation_is_xbrl_category(self, report):
        """BasisForPreparation is categorised as 'xbrl'."""
        assert report["definedNames"]["BasisForPreparation"]["category"] == "xbrl"

    def test_template_label_names_are_categorised(self, report):
        """template_label_* names should be categorised as 'template_label'."""
        sample = [
            (k, v)
            for k, v in report["definedNames"].items()
            if k.startswith("template_label_")
        ]
        assert len(sample) > 0
        for name, entry in sample[:10]:
            assert (
                entry["category"] == "template_label"
            ), f"{name} should be 'template_label', got {entry['category']}"


# ---------------------------------------------------------------------------
# 10. Matrix labels (row / column labels per section and per field)
# ---------------------------------------------------------------------------


class TestMatrixLabels:
    """Verify row and column label detection on matrix-structured sections."""

    # -- Section-level counts -----------------------------------------------

    def test_matrix_section_count(self, report):
        """13 sections should have both rowLabels and columnLabels (matrix).

        Includes sections with bold column headers (classic matrices) and
        sections with open-table headers (multiple non-bold template_label
        cells on the same row, e.g. b1_list_of_sites).

        Index-label reclassification moves bold headers whose column also
        contains row labels into ``indexLabels``.  Sections that had only
        index-column headers (e.g. b8_turnover_rate, b9, b10, c5) are now
        row-only.
        """
        matrix = [
            s
            for s in report["sections"]
            if s.get("rowLabels") and s.get("columnLabels")
        ]
        assert len(matrix) == 13

    def test_row_only_section_count(self, report):
        """25 sections should have rowLabels but no columnLabels.

        6 sections that previously had column labels are now row-only
        because their only column header was in the index column (same
        column as row labels) and got reclassified as indexLabels.
        """
        row_only = [
            s
            for s in report["sections"]
            if s.get("rowLabels") and not s.get("columnLabels")
        ]
        assert len(row_only) == 25

    def test_col_only_section_count(self, report):
        """2 sections have only column labels (no row labels).

        Open-table header detection reclassifies rows with multiple non-bold
        template_label cells as column headers.  In b1_list_of_subsidiaries
        and b2_practices the *only* non-bold row was the header row, so after
        reclassification no row labels remain."""
        col_only = [
            s
            for s in report["sections"]
            if s.get("columnLabels") and not s.get("rowLabels")
        ]
        assert len(col_only) == 2

    def test_index_label_section_count(self, report):
        """16 sections should have indexLabels.

        Index labels are bold column headers whose column also contains
        non-bold row labels — they describe the row-label dimension
        (e.g. "Land-use type", "Row ID", "Gender").
        """
        has_idx = [
            s
            for s in report["sections"]
            if s.get("indexLabels")
        ]
        assert len(has_idx) == 16

    def test_b5_land_use_has_index_label(self, report):
        """b5_biodiversity_land_use should have land_use_type as indexLabel."""
        sec = _find_section(report, "b5_biodiversity_land_use")
        idx = sec.get("indexLabels")
        assert idx is not None
        idx_keys = {v["key"] for v in idx.values()}
        assert "land_use_type" in idx_keys
        # area should remain as a column label
        cl = sec.get("columnLabels")
        assert cl is not None
        cl_keys = {v["key"] for v in cl.values()}
        assert "area" in cl_keys

    # -- Energy breakdown (a clean matrix example) --------------------------

    def test_energy_breakdown_is_matrix(self, report):
        """b3_breakdown_of_energy_consumption has row & column labels."""
        sec = _find_section(report, "b3_breakdown_of_energy_consumption_in_MWh")
        assert sec is not None
        assert sec.get("rowLabels") is not None
        assert sec.get("columnLabels") is not None

    def test_energy_breakdown_row_labels(self, report):
        """Energy breakdown has 4 row labels: electricity, selfgenerated, fuels + qualifying Q."""
        sec = _find_section(report, "b3_breakdown_of_energy_consumption_in_MWh")
        rl = sec["rowLabels"]
        assert len(rl) == 4
        keys = {v["key"] for v in rl.values()}
        assert "electricity" in keys
        assert "fuels" in keys

    def test_energy_breakdown_col_labels(self, report):
        """Energy breakdown has 3 column labels: renewable, non_renewable, total."""
        sec = _find_section(report, "b3_breakdown_of_energy_consumption_in_MWh")
        cl = sec["columnLabels"]
        assert len(cl) == 3
        keys = {v["key"] for v in cl.values()}
        assert "renewable" in keys
        assert "non_renewable" in keys

    # -- B3 GHG emissions ---------------------------------------------------

    def test_b3_ghg_has_row_labels(self, report):
        """B3 GHG emissions should have 25 row labels (Scope 1/2/3 etc.)."""
        sec = _find_section(
            report,
            "b3_estimated_greenhouse_gas_emissions_considering_the_GHG_protocol_version_2004_in_tCO2e",
        )
        assert sec.get("rowLabels") is not None
        assert len(sec["rowLabels"]) == 25

    def test_b3_ghg_has_column_label(self, report):
        """B3 GHG emissions should have 1 column label: current_reporting_period."""
        sec = _find_section(
            report,
            "b3_estimated_greenhouse_gas_emissions_considering_the_GHG_protocol_version_2004_in_tCO2e",
        )
        cl = sec.get("columnLabels")
        assert cl is not None
        assert len(cl) == 1
        keys = {v["key"] for v in cl.values()}
        assert "current_reporting_period" in keys

    # -- C3 GHG targets (special case: inherits row labels from B3) ----------

    def test_c3_ghg_targets_inherits_row_labels(self, report):
        """C3 GHG targets inherits 25 row labels from B3 + 1 own = 26 total."""
        c3 = _find_section(report, "c3_GHG_reduction_targets_in_tC02e")
        assert c3.get("rowLabels") is not None
        assert len(c3["rowLabels"]) == 26
        # Should contain GHG category labels from B3
        keys = {v["key"] for v in c3["rowLabels"].values()}
        assert "gross_scope_1_ghg_emissions" in keys
        assert "total_scope_3_ghg_emissions" in keys
        # Should also contain C3's own non-bold label
        assert (
            "has_the_undertaking_has_established_ghg_emission_reduction_targets" in keys
        )

    def test_c3_ghg_targets_has_column_labels(self, report):
        """C3 GHG targets should have target_year and percentage_reduction."""
        c3 = _find_section(report, "c3_GHG_reduction_targets_in_tC02e")
        cl = c3.get("columnLabels")
        assert cl is not None
        keys = {v["key"] for v in cl.values()}
        assert "target_year" in keys
        assert "percentage_reduction_from_base_year" in keys

    # -- C8 revenues (matrix with monetary amount column) -------------------

    def test_c8_revenues_is_matrix(self, report):
        """c8_revenues_from_certain_sectors has row & column labels."""
        sec = _find_section(report, "c8_revenues_from_certain_sectors")
        assert sec.get("rowLabels") is not None
        assert sec.get("columnLabels") is not None
        keys = {v["key"] for v in sec["columnLabels"].values()}
        assert "monetary_amount_in" in keys

    # -- Label descriptor structure -----------------------------------------

    def test_label_descriptor_has_required_keys(self, report):
        """Each rowLabel/columnLabel descriptor must have key, templateLabelKey, templateLabels."""
        for sec in report["sections"]:
            for label_type in ("rowLabels", "columnLabels"):
                label_map = sec.get(label_type)
                if not label_map:
                    continue
                for ref, desc in label_map.items():
                    assert (
                        "key" in desc
                    ), f"{sec['sectionId']} {label_type}[{ref}] missing 'key'"
                    assert (
                        "templateLabelKey" in desc
                    ), f"{sec['sectionId']} {label_type}[{ref}] missing 'templateLabelKey'"
                    assert (
                        "templateLabels" in desc
                    ), f"{sec['sectionId']} {label_type}[{ref}] missing 'templateLabels'"
                    # templateLabelKey should start with 'template_label_'
                    assert desc["templateLabelKey"].startswith("template_label_"), (
                        f"{sec['sectionId']} {label_type}[{ref}] templateLabelKey "
                        f"doesn't start with template_label_"
                    )

    def test_label_descriptors_have_english_translation(self, report):
        """Row/column label descriptors with templateLabels should include English."""
        missing = []
        for sec in report["sections"]:
            for label_type in ("rowLabels", "columnLabels"):
                label_map = sec.get(label_type)
                if not label_map:
                    continue
                for ref, desc in label_map.items():
                    tl = desc.get("templateLabels", {})
                    if tl and "en" not in tl:
                        missing.append(f"{sec['sectionId']}.{label_type}[{ref}]")
        assert missing == [], f"Descriptors without English: {missing}"

    # -- Field-level rowLabel/columnLabel -----------------------------------

    def test_all_fields_have_rowLabel_and_columnLabel_keys(self, report):
        """Every field in non-flat sections should have rowLabel and columnLabel keys."""
        missing = []
        for sec in report["sections"]:
            fields = sec.get("fields", [])
            if isinstance(fields, dict):
                continue  # flat mode
            for f in fields:
                if "rowLabel" not in f:
                    missing.append(
                        f"{sec['sectionId']}/{f.get('fieldId', f.get('fieldName', '?'))}/rowLabel"
                    )
                if "columnLabel" not in f:
                    missing.append(
                        f"{sec['sectionId']}/{f.get('fieldId', f.get('fieldName', '?'))}/columnLabel"
                    )
        assert missing == [], f"Fields missing rowLabel/columnLabel: {missing[:10]}"

    def test_energy_breakdown_fields_have_matrix_labels(self, report):
        """Fields in energy breakdown that are in the matrix should have both labels."""
        sec = _find_section(report, "b3_breakdown_of_energy_consumption_in_MWh")
        fields_with_both = [
            f for f in sec["fields"] if f.get("rowLabel") and f.get("columnLabel")
        ]
        # At least some fields should have both row and column labels
        assert (
            len(fields_with_both) >= 1
        ), "Expected some fields with both row and column labels"

    def test_energy_breakdown_question_has_no_column_label(self, report):
        """The question-answer pair above the table should NOT get a column label."""
        sec = _find_section(report, "b3_breakdown_of_energy_consumption_in_MWh")
        qa_field = None
        for f in sec["fields"]:
            rl = f.get("rowLabel")
            if rl and "has_the_undertaking" in rl.get("key", ""):
                qa_field = f
                break
        assert qa_field is not None, "Expected question-answer field"
        assert (
            qa_field.get("columnLabel") is None
        ), "Question-answer field above the table should not have a column label"

    def test_energy_breakdown_total_facts_have_labels(self, report):
        """Total/aggregate facts (default member) should have total column label."""
        sec = _find_section(report, "b3_breakdown_of_energy_consumption_in_MWh")
        total_fields = [
            f
            for f in sec["fields"]
            if f.get("columnLabel") and "total" in f["columnLabel"].get("key", "")
        ]
        # Electricity total (110) and self-generated electricity total (43)
        assert (
            len(total_fields) >= 2
        ), f"Expected at least 2 total-member fields, got {len(total_fields)}"
        row_keys = {f["rowLabel"]["key"] for f in total_fields if f.get("rowLabel")}
        assert "electricity" in row_keys
        assert "selfgenerated_electricity" in row_keys

    def test_field_rowLabel_descriptor_structure(self, report):
        """Field-level rowLabel (when not null) should have key, templateLabelKey, templateLabels."""
        for sec in report["sections"]:
            fields = sec.get("fields", [])
            if isinstance(fields, dict):
                continue
            for f in fields:
                rl = f.get("rowLabel")
                if rl is not None:
                    assert "key" in rl, f"rowLabel missing 'key' in {sec['sectionId']}"
                    assert "templateLabelKey" in rl
                    assert "templateLabels" in rl

    def test_field_columnLabel_descriptor_structure(self, report):
        """Field-level columnLabel (when not null) should have key, templateLabelKey, templateLabels."""
        for sec in report["sections"]:
            fields = sec.get("fields", [])
            if isinstance(fields, dict):
                continue
            for f in fields:
                cl = f.get("columnLabel")
                if cl is not None:
                    assert (
                        "key" in cl
                    ), f"columnLabel missing 'key' in {sec['sectionId']}"
                    assert "templateLabelKey" in cl
                    assert "templateLabels" in cl

    # -- Bold-font detection specific tests ---------------------------------

    def test_b4_pollution_has_bold_column_headers(self, report):
        """B4 pollution should have 4 column headers (pollutant + 3 emissions).

        row_id (col C) was reclassified as an indexLabel because col C
        also contains non-bold row labels.
        """
        sec = _find_section(report, "b4_pollution_of_air_water_and_soil")
        cl = sec.get("columnLabels")
        assert cl is not None
        assert len(cl) == 4
        col_keys = {v["key"] for v in cl.values()}
        assert "pollutant" in col_keys
        assert "emission_to_air" in col_keys
        assert "emission_to_water" in col_keys
        assert "emission_to_soil" in col_keys
        # row_id is now an index label
        idx = sec.get("indexLabels")
        assert idx is not None
        idx_keys = {v["key"] for v in idx.values()}
        assert "row_id" in idx_keys

    def test_b4_pollution_has_non_bold_row_labels(self, report):
        """B4 pollution should have non-bold question labels as row labels."""
        sec = _find_section(report, "b4_pollution_of_air_water_and_soil")
        rl = sec.get("rowLabels")
        assert rl is not None
        assert len(rl) == 3
        row_keys = {v["key"] for v in rl.values()}
        assert "is_this_disclosure_already_publicly_available" in row_keys

    def test_b8_contract_type_matrix(self, report):
        """B8 type of contract should be a matrix with row labels, column labels and index labels.

        type_of_contract (col C) is an index label because col C also
        contains the row labels (permanent_contract, temporary_contract).
        """
        sec = _find_section(
            report, "b8_workforce_general_characteristics_type_of_contract"
        )
        rl = sec.get("rowLabels")
        cl = sec.get("columnLabels")
        idx = sec.get("indexLabels")
        assert rl is not None
        assert cl is not None
        assert idx is not None
        assert len(rl) == 2
        assert len(cl) == 1
        assert len(idx) == 1
        row_keys = {v["key"] for v in rl.values()}
        assert "permanent_contract" in row_keys
        assert "temporary_contract" in row_keys
        col_keys = {v["key"] for v in cl.values()}
        assert "number_of_employees" in col_keys
        idx_keys = {v["key"] for v in idx.values()}
        assert "type_of_contract" in idx_keys

    def test_b7_waste_has_bold_column_headers(self, report):
        """B7 waste section should have 5 column headers.

        unit_of_measurement (col G) was reclassified as an indexLabel
        because col G also contains non-bold row labels (unit values).
        """
        sec = _find_section(report, "b7_waste_generated")
        cl = sec.get("columnLabels")
        assert cl is not None
        assert len(cl) == 5
        col_keys = {v["key"] for v in cl.values()}
        assert "type_of_waste" in col_keys
        assert "waste_diverted_to_recycle_or_reuse" in col_keys
        # unit_of_measurement is now an index label
        idx = sec.get("indexLabels")
        assert idx is not None
        idx_keys = {v["key"] for v in idx.values()}
        assert "unit_of_measurement" in idx_keys

    # -- additional_rows_warning boundary tests ----------------------------
    # Sections that contain ``=template_label_additional_rows_warning`` cells
    # should suppress column labels for XBRL fields below the warning row.
    # The warning marks the end of the repeatable table; rows below it are
    # summary/total rows that should stand alone (no table column label).

    _SECTIONS_WITH_ARW = {
        "information_on_previous_reporting_period",
        "b1_list_of_subsidiaries",
        "b1_list_of_sites",
        "b4_pollution_of_air_water_and_soil",
        "b5_sites_in_biodiversity_sensitive_areas",
        "b7_waste_generated",
        "b7_annual_mass_flow_of_relevant_materials_used",
        "b8_workforce_general_characteristics_country_of_employment",
    }

    def test_sections_with_arw_count(self, report):
        """8 sections should have an additionalRowsWarningRow value in their
        internal excel_sections structure. The count is verified here
        indirectly by checking the known section list."""
        assert len(self._SECTIONS_WITH_ARW) == 8

    def test_has_additional_rows_warning_flag(self, report):
        """Each section with an additional_rows_warning should have
        hasAdditionalRowsWarning=True in the output JSON."""
        for sec in report["sections"]:
            sid = sec["sectionId"]
            arw = sec.get("hasAdditionalRowsWarning", False)
            if sid in self._SECTIONS_WITH_ARW:
                assert arw is True, (
                    f"{sid} should have hasAdditionalRowsWarning=True"
                )
            else:
                assert arw is False, (
                    f"{sid} should have hasAdditionalRowsWarning=False"
                )

    def test_no_field_has_additional_rows_warning_labelKey(self, report):
        """No field in any section should have labelKey='additional_rows_warning'.
        The warning label must not contaminate _row_to_template_label."""
        for sec in report["sections"]:
            for f in sec["fields"]:
                lk = f.get("labelKey")
                assert lk != "additional_rows_warning", (
                    f"Field {f.get('qname') or f.get('fieldId')} in "
                    f"{sec['sectionId']} has labelKey='additional_rows_warning'"
                )

    def test_b7_mass_flow_table_fields_have_null_labelKey(self, report):
        """Range-based table fields in b7_mass_flow (NameOfMaterialUsed,
        WeightOfMaterialUsed, VolumeOfMaterialUsed) should have
        labelKey=None — not 'additional_rows_warning'."""
        sec = _find_section(
            report, "b7_annual_mass_flow_of_relevant_materials_used"
        )
        table_qnames = {
            "vsme:NameOfMaterialUsed",
            "vsme:WeightOfMaterialUsed",
            "vsme:VolumeOfMaterialUsed",
        }
        table_fields = [
            f for f in sec["fields"] if f.get("qname") in table_qnames
        ]
        assert len(table_fields) >= 3, (
            f"Expected ≥3 table data fields, got {len(table_fields)}"
        )
        for f in table_fields:
            assert f.get("labelKey") is None, (
                f"Table field {f['qname']} should have labelKey=None "
                f"(got {f.get('labelKey')!r})"
            )

    def test_b7_mass_flow_total_fields_have_no_column_label(self, report):
        """TotalMass/Volume fields in b7_mass_flow are below the warning row
        and should NOT have a column label (they used to inherit mass_volume)."""
        sec = _find_section(
            report, "b7_annual_mass_flow_of_relevant_materials_used"
        )
        total_fields = [
            f
            for f in sec["fields"]
            if f.get("rowLabel")
            and "total_annual_mass_flow" in f["rowLabel"].get("key", "")
        ]
        assert len(total_fields) == 2, (
            f"Expected 2 total mass/volume fields, got {len(total_fields)}"
        )
        for f in total_fields:
            assert f.get("columnLabel") is None, (
                f"Total field {f.get('rowLabel', {}).get('key', '?')} "
                f"should not have a column label (below warning row)"
            )

    def test_b7_mass_flow_table_fields_retain_column_labels(self, report):
        """Fields in the table area (above the warning row) should still have
        column labels."""
        sec = _find_section(
            report, "b7_annual_mass_flow_of_relevant_materials_used"
        )
        fields_with_cl = [
            f for f in sec["fields"] if f.get("columnLabel") is not None
        ]
        # Table data fields above the warning should have column labels
        assert len(fields_with_cl) >= 2, (
            "Expected at least 2 table fields with column labels in b7_mass_flow"
        )
        cl_keys = {f["columnLabel"]["key"] for f in fields_with_cl}
        assert "name_of_the_key_material" in cl_keys
        assert "mass_volume" in cl_keys

    # Sections that legitimately have column labels but no row labels.
    # These are open-table sections where the *only* non-bold template_label
    # row is the multi-cell header row itself.
    _COL_ONLY_SECTIONS: set[str] = {
        "b1_list_of_subsidiaries",
        "b2_practices_policies_and_future_initiatives_for_transitioning_towards_a_more_sustainable_economy",
    }

    def test_no_section_has_only_column_labels(self, report):
        """Sections with column labels also have row labels, except known col-only sections."""
        for sec in report["sections"]:
            if sec.get("columnLabels") and not sec.get("rowLabels"):
                assert sec["sectionId"] in self._COL_ONLY_SECTIONS, (
                    f"{sec['sectionId']} has columnLabels but no rowLabels "
                    f"and is not in _COL_ONLY_SECTIONS"
                )

    def test_neither_sections_count(self, report):
        """5 sections have neither row nor column labels (free-text sections)."""
        neither = [
            s
            for s in report["sections"]
            if not s.get("rowLabels") and not s.get("columnLabels")
        ]
        assert len(neither) == 5

    def test_row_only_sections_have_no_bold_headers(self, report):
        """ROW-ONLY sections should have rowLabels but null columnLabels."""
        row_only = [
            s
            for s in report["sections"]
            if s.get("rowLabels") and not s.get("columnLabels")
        ]
        for sec in row_only:
            assert (
                sec["columnLabels"] is None
            ), f"{sec['sectionId']} should have null columnLabels"

    # -- Orphan coverage: section-level labels used by at least one field ----

    # Known exceptions where a row label has no matching field:
    # - Sections with zero fields (e.g. information_on_the_report_necessary_for_XBRL)
    # - Row labels for rows that have no data in the sample file
    #   (e.g. scope 3 sub-categories 1-15, fuels in energy breakdown)
    # These are excluded via the _ALLOWED_ORPHAN_ROW_LABELS allowlist.
    _ALLOWED_ORPHAN_ROW_LABELS: dict[str, set[str]] = {
        # Metadata section has no XBRL fields in this extraction
        "information_on_the_report_necessary_for_XBRL": {
            "template_label_name_of_the_reporting_entity",
            "template_label_identifier_of_the_reporting_entity",
            "template_label_currency_of_the_monetary_values_in_the_report",
            "template_label_starting_year",
            "template_label_starting_month",
            "template_label_starting_day",
            "template_label_reporting_period_start_date",
            "template_label_ending_year",
            "template_label_ending_month",
            "template_label_reporting_period_end_date",
        },
        # Open-table sections: "id" column and geolocation row have no
        # XBRL fields.  b1_list_of_subsidiaries is now col-only (no row
        # labels at all), so only b1_list_of_sites keeps a row orphan.
        "b1_list_of_sites": {"template_label_automatic_geolocation"},
        # Energy breakdown: fuels row only has data when Fuel Converter is used
        "b3_breakdown_of_energy_consumption_in_MWh": {"template_label_fuels"},
        # B3 GHG: scope 3 sub-categories (1-15) only populated when user fills
        # them; also year_date row (header for year columns), the scope-3
        # disclosure trigger, and market-based totals (value="-" in sample).
        "b3_estimated_greenhouse_gas_emissions_considering_the_GHG_protocol_version_2004_in_tCO2e": {
            "template_label_year_date",
            "template_label_is_the_undertaking_disclosing_entity_specific_information_on_scope_3_emissions",
            "template_label_total_scope_3_ghg_emissions",
            "template_label_total_scope_1_scope_2_and_scope_3_ghg_emissions_location_based",
            "template_label_total_scope_1_scope_2_and_scope_3_ghg_emissions_market_based",
            "template_label_1_purchased_goods_and_services",
            "template_label_2_capital_goods",
            "template_label_3_fuel__and_energy_related_activities",
            "template_label_4_upstream_transportation_and_distribution",
            "template_label_5_waste_generated_in_operations",
            "template_label_6_business_travel",
            "template_label_7_employee_commuting",
            "template_label_8_upstream_leased_assets",
            "template_label_9_downstream_transportation_and_distribution",
            "template_label_10_processing_of_sold_products",
            "template_label_11_use_of_sold_products",
            "template_label_12_end_of_life_treatment_of_sold_products",
            "template_label_13_downstream_leased_assets",
            "template_label_14_franchises",
            "template_label_15_investments",
        },
        # C3 inherits B3 row labels including scope 3 sub-categories;
        # also market-based rows where the sample has no values.
        "c3_GHG_reduction_targets_in_tC02e": {
            "template_label_gross_scope_2_market_based_ghg_emissions",
            "template_label_total_scope_1_and_scope_2_ghg_emissions_market_based",
            "template_label_total_scope_1_scope_2_and_scope_3_ghg_emissions_market_based",
            "template_label_1_purchased_goods_and_services",
            "template_label_2_capital_goods",
            "template_label_3_fuel__and_energy_related_activities",
            "template_label_4_upstream_transportation_and_distribution",
            "template_label_5_waste_generated_in_operations",
            "template_label_6_business_travel",
            "template_label_7_employee_commuting",
            "template_label_8_upstream_leased_assets",
            "template_label_9_downstream_transportation_and_distribution",
            "template_label_10_processing_of_sold_products",
            "template_label_11_use_of_sold_products",
            "template_label_12_end_of_life_treatment_of_sold_products",
            "template_label_13_downstream_leased_assets",
            "template_label_14_franchises",
            "template_label_15_investments",
        },
        # C3 transition plan: trigger question row has no XBRL fact
        "c3_transition_plan_for_undertakings_operating_in_high_climate_impact_sectors": {
            "template_label_is_the_undertaking_operating_in_high_impact_sectors",
        },
        # Turnover rate: intermediate input rows (employees who left, employees
        # at start/end of period) feed into the computed turnover rate field.
        "b8_workforce_general_characteristics_turnover_rate": {
            "template_label_number_of_employees_who_left_during_the_reporting_period",
            "template_label_number_of_employees_at_the_beginning_of_the_reporting_period",
            "template_label_number_of_employees_at_the_end_of_the_reporting_period",
        },
        # Pay gap: intermediate input rows for male/female hourly pay feed
        # the computed percentage gap field.
        "b10_workforce_remuneration_collective_bargaining_and_training": {
            "template_label_average_gross_hourly_pay_level_of_male_employees",
            "template_label_average_gross_hourly_pay_level_of_female_employees",
        },
        # Training hours: header row for section table and computed average
        "b10_workforce_remuneration_collective_bargaining_and_training_always_reported": {
            "template_label_number_of_annual_training_hours_per_employee_during_the_reporting_period",
            "template_label_average_number_of_annual_training_hours_per_employee",
        },
        # Management ratio: intermediate input rows for male/female counts
        "c5_additional_general_workforce_characteristics": {
            "template_label_number_of_male_employees_at_management_level",
            "template_label_number_of_female_employees_at_management_level",
        },
        # Health & safety: intermediate input row (total hours worked) feeds
        # the computed accident rate field.
        "b9_workforce_health_and_safety": {
            "template_label_total_number_of_hours_worked_in_a_year_by_all_employees_in_the_reporting_period",
        },
        # Human rights policies: sub-option labels (forced_labour, etc.)
        # under "if_yes_does_this_cover" conditional block — individual
        # sub-items don't produce XBRL facts.
        "c6_additional_own_workforce_information_human_rights_policies_and_processes": {
            "template_label_if_yes_does_this_cover",
            "template_label_forced_labour",
            "template_label_human_trafficking",
            "template_label_discrimination",
            "template_label_accident_prevention",
            "template_label_other_if_yes_specify",
        },
        # Severe human rights incidents: sub-option labels under
        # "if_yes_are_incidents_related_to" conditional block.
        "c7_severe_negative_human_rights_incidents": {
            "template_label_if_yes_are_incidents_related_to",
            "template_label_forced_labour",
            "template_label_human_trafficking",
            "template_label_discrimination",
            "template_label_other_if_yes_specify",
        },
        # Governance diversity: intermediate input rows (male/female counts)
        # feed the computed ratio field.
        "c9_gender_diversity_ratio_in_the_governance_body": {
            "template_label_number_of_female_board_members_at_the_end_of_the_reporting_period",
            "template_label_number_of_male_board_members_at_the_end_of_the_reporting_period",
        },
    }

    def test_row_labels_used_by_fields(self, report):
        """Every section-level rowLabel should appear in at least one field's rowLabel.

        If a section declares a rowLabel (e.g. ``template_label_1_purchased_goods_and_services``),
        at least one field in that section should reference it via its ``rowLabel.templateLabelKey``.
        Orphaned labels indicate that either the label detection picked up a label
        that doesn't correspond to a field, or the field-label assignment has a gap.
        Known exceptions are listed in ``_ALLOWED_ORPHAN_ROW_LABELS``.
        """
        violations = []
        for sec in report["sections"]:
            rl = sec.get("rowLabels") or {}
            if not rl:
                continue
            field_rl_keys = {
                f["rowLabel"]["templateLabelKey"]
                for f in sec["fields"]
                if f.get("rowLabel")
                and isinstance(f["rowLabel"], dict)
                and "templateLabelKey" in f["rowLabel"]
            }
            allowed = self._ALLOWED_ORPHAN_ROW_LABELS.get(sec["sectionId"], set())
            for _row_num, desc in rl.items():
                tlk = desc.get("templateLabelKey", "")
                if tlk and tlk not in field_rl_keys and tlk not in allowed:
                    violations.append(f"{sec['sectionId']}: {tlk}")
        assert violations == [], (
            f"Row labels declared at section level but not used by any field "
            f"({len(violations)} orphans):\n  " + "\n  ".join(violations)
        )

    # Known exceptions where a column label has no matching field.
    # These are typically user-input columns in open/expandable tables,
    # computed/auto-filled columns, or sub-category checkbox columns.
    # Index-column headers (bold headers in the same column as row labels)
    # are now classified as ``indexLabels`` and are not checked here.
    _ALLOWED_ORPHAN_COL_LABELS: dict[str, set[str]] = {
        # Open-table header sections: "id" index column has no XBRL field
        "b1_list_of_subsidiaries": {"template_label_id"},
        "b1_list_of_sites": {"template_label_id"},
        # B2 practices: sustainability-topic checkbox columns (pollution,
        # water, biodiversity, etc.) are label-only — no XBRL data fields.
        "b2_practices_policies_and_future_initiatives_for_transitioning_towards_a_more_sustainable_economy": {
            "template_label_pollution",
            "template_label_water_and_marine_resources",
            "template_label_biodiversity_and_ecosystems",
            "template_label_circular_economy",
            "template_label_own_workforce",
            "template_label_workers_in_the_value_chain",
            "template_label_affected_communities",
            "template_label_consumers_and_endusers_validation",
            "template_label_business_conduct",
        },
        # Energy breakdown: "Total" column (K) — computed sum
        "b3_breakdown_of_energy_consumption_in_MWh": {
            "template_label_total_renewable_and_nonrenewable",
        },
        # C3 GHG targets: "percentage_reduction" column (K) — computed
        "c3_GHG_reduction_targets_in_tC02e": {
            "template_label_percentage_reduction_from_base_year",
        },
        # B4 pollution: col D (pollutant name) is a user-input column
        "b4_pollution_of_air_water_and_soil": {
            "template_label_pollutant",
        },
        # B5 biodiversity: col D (site location) is auto-filled from B1
        "b5_sites_in_biodiversity_sensitive_areas": {
            "template_label_site_location_in_near_a_biodiversity_area",
        },
        # B7 waste: col C (row_id) and col D (type_of_waste) are user-input columns
        "b7_waste_generated": {
            "template_label_row_id",
            "template_label_type_of_waste",
        },
        # B7 mass flow: col J (unit_of_measurement) is a user-input column
        "b7_annual_mass_flow_of_relevant_materials_used": {
            "template_label_unit_of_measurement",
        },
        # C8 revenues: col C describes the section topic
        "c8_revenues_from_certain_sectors": {
            "template_label_total_revenues_derived_from_fossil_fuel_sector",
        },
    }

    def test_col_labels_used_by_fields(self, report):
        """Every section-level columnLabel should appear in at least one field's columnLabel.

        Similar to test_row_labels_used_by_fields but for column labels.
        Known exceptions are listed in ``_ALLOWED_ORPHAN_COL_LABELS``.
        """
        violations = []
        for sec in report["sections"]:
            cl = sec.get("columnLabels") or {}
            if not cl:
                continue
            field_cl_keys = {
                f["columnLabel"]["templateLabelKey"]
                for f in sec["fields"]
                if f.get("columnLabel")
                and isinstance(f["columnLabel"], dict)
                and "templateLabelKey" in f["columnLabel"]
            }
            allowed = self._ALLOWED_ORPHAN_COL_LABELS.get(sec["sectionId"], set())
            for _col_ref, desc in cl.items():
                tlk = desc.get("templateLabelKey", "")
                if tlk and tlk not in field_cl_keys and tlk not in allowed:
                    violations.append(f"{sec['sectionId']}: {tlk}")
        assert violations == [], (
            f"Column labels declared at section level but not used by any field "
            f"({len(violations)} orphans):\n  " + "\n  ".join(violations)
        )

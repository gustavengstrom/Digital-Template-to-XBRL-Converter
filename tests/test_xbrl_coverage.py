"""
XBRL coverage tests for report-facts.json.

These tests verify that every ``definedNames`` entry with ``"category": "xbrl"``
is represented as a field somewhere in the ``sections`` output.

Run with:
    conda run -n p312 python -m pytest tests/test_xbrl_coverage.py -v

Background
----------
``definedNames`` in the report-facts.json maps XBRL concept identifiers to cell
references in the Excel template.  XBRL-category entries fall into several
sub-categories:

Structural (not data fields — excluded from coverage checks):
  - ``*_unit``     – unit references (e.g. ``TotalEnergyConsumption_unit``)
  - ``*Table``     – XBRL table containers
  - ``*Axis``      – dimensional axes
  - ``*Hypercube`` – hypercube definitions
  - Standalone ``*Member`` names (no underscore) – e.g. ``BaselineYearMember``

Data concepts (the ones we check):
  - Simple concepts: ``TotalEnergyConsumption`` → matches field with
    ``qname == "vsme:TotalEnergyConsumption"``
  - Dimensioned concepts: ``EnergyConsumptionFromElectricity_NonRenewableEnergyMember``
    → matches field with ``qname == "vsme:EnergyConsumptionFromElectricity"`` AND
    a dimension member ``vsme:NonRenewableEnergyMember``
  - Range-based dynamic concepts: ``AmountOfEmissionToAir`` with cellRef
    ``'Environmental Disclosures'!$G$80:$I$179`` → matches fields that share the
    same base qname but have varying typed-dimension members
"""

import json
import re
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
def defined_names(report) -> dict[str, Any]:
    return report.get("definedNames", {})


@pytest.fixture(scope="session")
def sections(report) -> list[dict]:
    return report.get("sections", [])


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _get_xbrl_defined_names(defined_names: dict) -> dict[str, Any]:
    """Return only the XBRL-category definedNames."""
    return {k: v for k, v in defined_names.items() if v.get("category") == "xbrl"}


def _get_xbrl_structural_names(xbrl_names: dict) -> dict[str, Any]:
    """Return structural XBRL names (not data fields): _unit, Table, Axis,
    Hypercube, standalone Member."""
    result = {}
    for name, info in xbrl_names.items():
        if (
            name.endswith("_unit")
            or name.endswith("Table")
            or name.endswith("Axis")
            or name.endswith("Hypercube")
            or (name.endswith("Member") and "_" not in name)
        ):
            result[name] = info
    return result


def _get_xbrl_data_names(xbrl_names: dict) -> dict[str, Any]:
    """Return XBRL data concept names (excluding structural names)."""
    structural = _get_xbrl_structural_names(xbrl_names)
    return {k: v for k, v in xbrl_names.items() if k not in structural}


def _build_field_index(sections: list[dict]) -> set[tuple[str, frozenset]]:
    """Build a set of (concept_local, frozenset((axis_local, member_local), ...))
    tuples from all section fields."""
    index: set[tuple[str, frozenset]] = set()
    for section in sections:
        for field in section.get("fields", []):
            qname = field.get("qname", "")
            if not qname or ":" not in qname:
                continue
            local = qname.split(":", 1)[1]
            dims = field.get("dimensions", {})
            dim_tuples = frozenset(
                (
                    a.split(":", 1)[1] if ":" in a else a,
                    m.split(":", 1)[1] if ":" in m else m,
                )
                for a, m in dims.items()
            )
            index.add((local, dim_tuples))
    return index


def _build_concept_set(sections: list[dict]) -> set[str]:
    """Build a set of all qname local parts (ignoring dimensions)."""
    concepts: set[str] = set()
    for section in sections:
        for field in section.get("fields", []):
            qname = field.get("qname", "")
            if qname and ":" in qname:
                concepts.add(qname.split(":", 1)[1])
    return concepts


def _match_defined_name(
    name: str,
    field_index: set[tuple[str, frozenset]],
    concept_set: set[str],
) -> str | None:
    """Try to match a defined name against the field index.

    Returns a description of the match type, or None if no match found.
    """
    # 1. Exact match as concept with no dimensions
    if (name, frozenset()) in field_index:
        return "exact_no_dims"

    # 2. Exact match as concept with any dimensions
    for concept, dims in field_index:
        if concept == name and dims:
            return "exact_with_dims"

    # 3. Compound name: concept + member suffix
    #    e.g. EnergyConsumptionFromElectricity_NonRenewableEnergyMember
    parts = name.split("_")
    for i in range(len(parts) - 1, 0, -1):
        candidate_concept = "_".join(parts[:i])
        candidate_member = "_".join(parts[i:])
        for concept, dims in field_index:
            if concept == candidate_concept:
                for _axis, member in dims:
                    if member == candidate_member:
                        return f"compound: concept={candidate_concept}, member={candidate_member}"
    return None


def _is_range_cellref(cellref: str) -> bool:
    """Check if a cellRef is a range (e.g. '$G$80:$I$179')."""
    if "!" not in cellref:
        return False
    cell_part = cellref.split("!")[-1]
    return ":" in cell_part


# ---------------------------------------------------------------------------
# Known gaps: XBRL defined names not represented in report-facts.json fields
# ---------------------------------------------------------------------------

# These are XBRL defined names that are legitimately NOT present as section
# fields due to how the Excel→JSON parser handles them.  Each is documented
# with a reason.
#
# When the parser is updated to cover these, remove them from this set and
# the test will start enforcing their presence.

KNOWN_MISSING: dict[str, str] = {
    # --- Energy breakdown: 'Total renewable and non-renewable' column ---
    # The 'Total' column in the energy breakdown table maps to fields with
    # dims={} (no dimension member), but the XBRL defined name expects the
    # TotalRenewableAndNonRenewableEnergyMember dimension.  The parser
    # currently emits these fields without that member.
    "EnergyConsumptionFromElectricity_TotalRenewableAndNonRenewableEnergyMember": (
        "Total column field has dims={} instead of TotalRenewableAndNonRenewableEnergyMember"
    ),
    "EnergyConsumptionFromSelfGeneratedElectricity_TotalRenewableAndNonRenewableEnergyMember": (
        "Total column field has dims={} instead of TotalRenewableAndNonRenewableEnergyMember"
    ),
    # --- EnergyConsumptionFromFuels: not parsed at all ---
    # Row 14 ('Fuels') in the energy breakdown table is not emitted as fields
    # by the parser, likely because the sample template has '-' values.
    "EnergyConsumptionFromFuels_RenewableEnergyMember": (
        "EnergyConsumptionFromFuels row not parsed (dash values)"
    ),
    "EnergyConsumptionFromFuels_NonRenewableEnergyMember": (
        "EnergyConsumptionFromFuels row not parsed (dash values)"
    ),
    "EnergyConsumptionFromFuels_TotalRenewableAndNonRenewableEnergyMember": (
        "EnergyConsumptionFromFuels row not parsed (dash values)"
    ),
    # --- GHG: b3 section missing dimension members ---
    # The b3 GHG section only has cl_key=current_reporting_period with dims={}.
    # XBRL expects CurrentlyStatedMember, but the parser does not assign it.
    # BaselineYear/TargetYear members only appear in c3 (reduction targets),
    # not in b3 for market-based scope 2 and market-based totals.
    "GrossScope3GreenhouseGasEmissions_CurrentlyStatedMember": (
        "b3 GHG section: CurrentlyStatedMember not in dims (only current_reporting_period cl_key)"
    ),
    "TotalGrossLocationBasedGHGEmissions_CurrentlyStatedMember": (
        "b3 GHG section: CurrentlyStatedMember not in dims (only in c3 with BaselineYear/TargetYear)"
    ),
    "GrossMarketBasedScope2GreenhouseGasEmissions_BaselineYearMember": (
        "Market-based scope 2 only in b3 (no BaselineYear dim); not in c3"
    ),
    "GrossMarketBasedScope2GreenhouseGasEmissions_TargetYearMember": (
        "Market-based scope 2 only in b3 (no TargetYear dim); not in c3"
    ),
    "TotalGrossMarketBasedGHGEmissions_BaselineYearMember": (
        "Market-based GHG total only in b3 (no BaselineYear dim); not in c3"
    ),
    "TotalGrossMarketBasedGHGEmissions_CurrentlyStatedMember": (
        "Market-based GHG total only in b3 (no CurrentlyStated dim); not in c3"
    ),
    "TotalGrossMarketBasedGHGEmissions_TargetYearMember": (
        "Market-based GHG total only in b3 (no TargetYear dim); not in c3"
    ),
    "TotalGrossMarketBasedScope1AndScope2GHGEmissions_BaselineYearMember": (
        "Market-based S1+S2 total only in b3 (no BaselineYear dim); not in c3"
    ),
    "TotalGrossMarketBasedScope1AndScope2GHGEmissions_TargetYearMember": (
        "Market-based S1+S2 total only in b3 (no TargetYear dim); not in c3"
    ),
}


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestXbrlDefinedNameCounts:
    """Verify expected counts of XBRL defined name categories."""

    def test_total_xbrl_defined_names(self, defined_names):
        """Total number of XBRL-category definedNames."""
        xbrl = _get_xbrl_defined_names(defined_names)
        assert len(xbrl) == 207

    def test_unit_names_count(self, defined_names):
        xbrl = _get_xbrl_defined_names(defined_names)
        units = {k for k in xbrl if k.endswith("_unit")}
        assert len(units) == 24

    def test_table_names_count(self, defined_names):
        xbrl = _get_xbrl_defined_names(defined_names)
        tables = {k for k in xbrl if k.endswith("Table")}
        assert len(tables) == 7

    def test_axis_names_count(self, defined_names):
        xbrl = _get_xbrl_defined_names(defined_names)
        axes = {k for k in xbrl if k.endswith("Axis")}
        assert len(axes) == 8

    def test_data_names_count(self, defined_names):
        """Data concept names (excluding structural names)."""
        xbrl = _get_xbrl_defined_names(defined_names)
        data = _get_xbrl_data_names(xbrl)
        assert len(data) == 165


class TestXbrlFieldCoverage:
    """Verify every XBRL data concept defined name maps to at least one
    section field."""

    def test_all_xbrl_data_names_have_fields(self, defined_names, sections):
        """Every XBRL data defined name must match a field (exact, compound,
        or range-based), or be listed in KNOWN_MISSING."""
        xbrl = _get_xbrl_defined_names(defined_names)
        data_names = _get_xbrl_data_names(xbrl)
        field_index = _build_field_index(sections)
        concept_set = _build_concept_set(sections)

        truly_missing = []
        for name, info in sorted(data_names.items()):
            if name in KNOWN_MISSING:
                continue

            match = _match_defined_name(name, field_index, concept_set)
            if match is not None:
                continue

            # For range-based cellRefs, check if the base concept exists
            # (dynamic tables have fields with typed dimensions)
            cellref = info.get("cellRef", "")
            if _is_range_cellref(cellref) and name in concept_set:
                continue

            truly_missing.append(name)

        if truly_missing:
            detail = "\n".join(
                f"  {n}: cellRef={data_names[n].get('cellRef')}" for n in truly_missing
            )
            pytest.fail(
                f"{len(truly_missing)} XBRL data defined name(s) have no "
                f"matching field in sections:\n{detail}"
            )

    def test_known_missing_count(self, defined_names):
        """Guard-rail: the number of known-missing items should not grow
        without deliberate change."""
        assert len(KNOWN_MISSING) == 14

    def test_known_missing_are_actually_missing(self, defined_names, sections):
        """Every entry in KNOWN_MISSING should genuinely be unmatched.
        If the parser is fixed, the entry should be removed."""
        xbrl = _get_xbrl_defined_names(defined_names)
        data_names = _get_xbrl_data_names(xbrl)
        field_index = _build_field_index(sections)
        concept_set = _build_concept_set(sections)

        false_known_missing = []
        for name in KNOWN_MISSING:
            if name not in data_names:
                continue
            info = data_names[name]
            match = _match_defined_name(name, field_index, concept_set)
            if match is not None:
                false_known_missing.append((name, match))
            elif _is_range_cellref(info.get("cellRef", "")) and name in concept_set:
                false_known_missing.append((name, "range_concept_match"))

        if false_known_missing:
            detail = "\n".join(
                f"  {n}: matched as '{m}' — remove from KNOWN_MISSING"
                for n, m in false_known_missing
            )
            pytest.fail(
                f"{len(false_known_missing)} KNOWN_MISSING entries now match "
                f"fields (parser was fixed?):\n{detail}"
            )


class TestXbrlStructuralNames:
    """Verify structural XBRL names are correctly categorised."""

    def test_structural_names_are_complete(self, defined_names):
        """All structural categories should sum to the expected total."""
        xbrl = _get_xbrl_defined_names(defined_names)
        structural = _get_xbrl_structural_names(xbrl)
        # _unit(24) + Table(7) + Axis(8) + Hypercube(1) + standalone Member(2) = 42
        assert len(structural) == 42

    def test_all_unit_names_have_data_counterpart(self, defined_names):
        """Every *_unit defined name should have a corresponding data concept
        (strip _unit suffix)."""
        xbrl = _get_xbrl_defined_names(defined_names)
        data_names = _get_xbrl_data_names(xbrl)
        unit_names = {k for k in xbrl if k.endswith("_unit")}

        missing_counterparts = []
        for uname in sorted(unit_names):
            base = uname[: -len("_unit")]
            if base not in data_names:
                missing_counterparts.append(uname)

        if missing_counterparts:
            pytest.fail(
                f"{len(missing_counterparts)} _unit name(s) lack a data "
                f"counterpart: {missing_counterparts}"
            )


class TestFieldDimensionCoverage:
    """Verify that fields with dimensions correctly reflect the XBRL taxonomy."""

    def test_energy_breakdown_has_renewable_and_nonrenewable(self, sections):
        """The energy breakdown section should have fields with both
        RenewableEnergyMember and NonRenewableEnergyMember."""
        section = next(
            s
            for s in sections
            if s.get("templateName")
            == "template_b3_breakdown_of_energy_consumption_in_MWh"
        )
        members = set()
        for field in section.get("fields", []):
            for _axis, member in field.get("dimensions", {}).items():
                local = member.split(":", 1)[1] if ":" in member else member
                members.add(local)

        assert "RenewableEnergyMember" in members
        assert "NonRenewableEnergyMember" in members

    def test_c3_ghg_targets_has_baseline_and_target_year(self, sections):
        """The C3 GHG reduction targets section should have fields with
        BaselineYearMember and TargetYearMember."""
        section = next(
            s
            for s in sections
            if s.get("templateName") == "template_c3_GHG_reduction_targets_in_tC02e"
        )
        members = set()
        for field in section.get("fields", []):
            for _axis, member in field.get("dimensions", {}).items():
                local = member.split(":", 1)[1] if ":" in member else member
                members.add(local)

        assert "BaselineYearMember" in members
        assert "TargetYearMember" in members

    def test_pollution_section_has_pollutant_dimensions(self, sections):
        """The pollution section fields should use the TypeOfPollutantAxis."""
        section = next(
            s
            for s in sections
            if s.get("templateName") == "template_b4_pollution_of_air_water_and_soil"
        )
        axes = set()
        for field in section.get("fields", []):
            for axis in field.get("dimensions", {}):
                local = axis.split(":", 1)[1] if ":" in axis else axis
                axes.add(local)

        assert "TypeOfPollutantAxis" in axes

    def test_waste_section_has_waste_type_dimensions(self, sections):
        """The waste section fields should use the TypeOfWasteAxis."""
        section = next(
            s
            for s in sections
            if s.get("templateName") == "template_b7_waste_generated"
        )
        axes = set()
        for field in section.get("fields", []):
            for axis in field.get("dimensions", {}):
                local = axis.split(":", 1)[1] if ":" in axis else axis
                axes.add(local)

        assert "TypeOfWasteAxis" in axes

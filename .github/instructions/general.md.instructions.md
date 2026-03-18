---
applyTo: '**'
---

This project aims to generate a JSON structure for building an online survey tool from a VSME Digital Excel Template. The JSON structure should be generated from the VSME Digital Template and underlying code, which reflects the [VSME Recommendation](https://eur-lex.europa.eu/eli/reco/2025/1710/oj/eng) as published by the European Commission on 30 July 2025.  

To complete the VSME Digital Template, preparers need to fill in the datapoints in the following four Excel sheets: 

1. General information sheet which contains the information necessary for the generation of the XBRL report and the general disclosures in VSME Basic and Comprehensive Module (B1, B2 and C1 and C2). It is essential for the XBRL converter to work properly that the cells “necessary for the generation of the XBRL report” are completed properly. Failure to complete those cells will trigger fatal errors when using the XRBL converter;

2. Environmental Disclosures sheet which contains the environmental metrics from both the Basic and Comprehensive Modules; 

3. Social Disclosures sheet which contains the social metrics from both the Basic and Comprehensive Modules; 

4. Governance Disclosures sheet which contains the governance metrics from both the Basic and Comprehensive Modules.

The Excel template along with the Python repository was built with the intention to be converted to XBRL format and an html report based on the data inserted into the excel by users. 

The challenge is to generate a JSON structure for building an online survey tool from the VSME Digital Excel Template and the code in the repository. Since the intention of the code in the repository is to convert the Excel template to XBRL format, the code is not structured in a way that makes it easy to generate a JSON structure for building an online survey tool.

## Excel Template
- When testing use the following template: `digital-templates/VSME-Digital-Template-Sample-1.1.1-unlocked.xlsx`

## Parsing

The script `scripts/parse-and-json.py` takes care of the excel template parsing. The output can be found in `output/report-facts.json`.

The general process is describe in `scripts/parse-and-json.md`.

## Generate survey_data json files

The script `scripts/gen-survey-json.py` takes care of converting the `output/report-facts.json` structure into a set of survey json file (one for each section). The survey json file structure (used for generating online surveys) is outlined in the `.github/instructions/survey_data.instructions.md` file.

## Notes:
### Validation
- The Excel template contains ~997 validation cells with IF formulas (referencing template_label_ok / template_label_missing_value) that indicate whether input fields are required, conditionally required, or optional. These were not being read by the converter.


### Sections
- Section headers have can be identified from defined names starting with `template_` 
(excluding those starting with `template_label`). Section headers also always have a fill color.
- Cells that contain values starting with `=template_label` used in the excel can be thought of as questions labels for a user answer given in a related cell. The associated answer cells are located to the right or below the question label cells. 
- Each section contains a sheet name. The script `scripts/gen-survey-json.py` should create a survey_group_name field in the output JSON structure for each section, with the value of the sheet name of the section. This field can be used to group sections by sheet in the survey tool.
- The output JSON structure should also include `sort_index` field which indicates the order of the sections and which order the survery should be accessed in the online survey tool.  

#### Section data
Section data can be structured within a section in several different ways. A single section may further contain mored than one type of data structure. Below are some examples of how the data can be structured:
- In some sections the questions labels are found in the leftmost column of the section, with answers in the adjacent column(s) to the right.
- In some sections, questions and answers are sometimes contained in a table. These tables have a header row where each column header has a value staring with `=template_label`. The column headers also aways use font weight bold (i.e. openpyxl: cell.font.bold = True). The header/column labels thus need to appear in the relevant field items of the section JSON structure output. An example is contained in section `template_b4_pollution_of_air_water_and_soil`.
- In some sections, questions and answers are sometimes contained in a table containing both a column and an index label. These tables have a header row where each column header has a value staring with `=template_label`. The column headers also aways use font weight bold (i.e. openpyxl: cell.font.bold = True). These tables also have an index column which is located to the left and contains cells with a value staring with `=template_label`. The header/column labels question labels and index colum label both need to to appear in the relevant field items of the section JSON structure output. An example is contained in section `template_b3_breakdown_of_energy_consumption_in_MWh`.
- Section `template_b3_estimated_greenhouse_gas_emissions_considering_the_GHG_protocol_version_2004_in_tCO2e` is a special case. It contains two tables of a table with multiple index labels (column C) and a shared column header (`template_label_current_reporting_period`). The first table follows directly below the header and ends with row 26. Cell C29 contains a question label with a boolean answer in Cell G29.
 The second table starts on row 30. 
- Section `template_c3_GHG_reduction_targets_in_tC02e` is a special case. It contains two tables of a table with multiple index labels (column C) and a shared column header (`template_label_current_reporting_period`). The first table follows directly below the header and ends with row 26. Cell C29 contains a question label with a boolean answer in Cell G29.
 The second table starts on row 30.  This section does however not contain the row labels in the leftmost column of the same section but intead uses the same row labels as used by section `template_b3_estimated_greenhouse_gas_emissions_considering_the_GHG_protocol_version_2004_in_tCO2e` (these sections are placed side-by-side in the excel). I also include a question - boolean answer pair on row 19 

# Notes on cell colors
- Section headers have a fill color (openpyxl: cell.fill.start_color) that can be used to identify them. The section header color is not the same across all sections but varies between sections.
- Cells with `cell.fill.start_color.rgb = "FFFFFF99"` (yellow) fill color are helper/condition fields that are not part of the XBRL report but facilitate completion of the disclosures. They include boolean trigger questions ("Has the undertaking…?"), helper numeric inputs, and dropdown selections. These cells can be identified by their fill color and the presence of a user-editable value (boolean, number, or dropdown). They also have a `template_label_*` formula on the same row that can be used to derive a stable identifier for them. These cells should be included in section field lists in the output JSON structure.
- Cells with `cell.fill.start_color.tint = 0.7999816888943144`: Indicates that the cell is automatically calculated and that no data entry is required. When converted to survey JSON, these cells should be marked with answer_type "COMPUTED" and not require user input.


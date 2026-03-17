# Survey Data Structure & Online Survey Generation

This document describes the JSON structure used for `survey_data_proxy` / `survey_data_snapshot` / `survey_data`, and how that structure drives the rendering of online (respondent-facing) surveys.

---

## Overview

The system uses a **three-layer JSON data model**:

| Layer | Model field | Owner | Purpose |
|---|---|---|---|
| **Survey definition** | `SurveyVersion.survey_data_snapshot` (surveybuilder) → `Survey.survey_data_proxy` (webhooks) | Admins | Master question list with full metadata |
| **Ticket responses** | `Ticket.survey_data` | Respondents | User-submitted answers, keyed by question ID |
| **Merged view** | Runtime result of `get_latest_survey_data_version()` | Home app | Full question + answer objects, used for display and reporting |

---

## 1. Survey Definition JSON (`survey_data_proxy`)

`Survey.survey_data_proxy` is a **flat list of question objects**. Each question object is produced by `Question.to_json_dict()` in `surveybuilder/models.py`.

### Minimal question object (all types)

```json
{
  "id": "ESRS_E1_abc12",
  "name": "What is your primary energy source?",
  "value": "",
  "answer_type": "RADIO",
  "help_text": "Select the most relevant option.",
  "related_ids": []
}
```

| Field | Type | Description |
|---|---|---|
| `id` | `string` | Unique question identifier within the survey. Format: `{survey_name}_{5_char_alphanumeric}`. Dynamic group questions append `__DQ0` suffix (e.g. `ESRS_E1_abc12__DQ0`). |
| `name` | `string` | The question text shown to the user. |
| `value` | `string` | Always `""` in the definition; populated with the respondent's answer in `Ticket.survey_data`. |
| `answer_type` | `string` | See [Answer Types](#2-answer-types) below. |
| `help_text` | `string` | Optional explanatory text shown below the question. |
| `related_ids` | `array[string]` | List of question IDs semantically related to this question (used for contextual display/reporting). |

---

### Optional fields (type-dependent)

#### Choice questions (`RADIO`, `DROPDOWN`, `CHECKBOX`)

```json
{
  "answer_options": [
    { "value": "coal", "label": "Coal" },
    { "value": "gas",  "label": "Natural Gas" },
    { "value": "solar","label": "Solar" }
  ]
}
```

`value` is the stored key; `label` is the display string. Labels can be updated via a patch without changing values.

#### Scale questions (`SCALE`)

```json
{
  "scale_min": 1,
  "scale_max": 5,
  "scale_step": 1
}
```

#### Reference choice questions (`RADIO_REF`, `DROPDOWN_REF`, `CHECKBOX_REF`)

```json
{
  "answer_options": { "id": "ESRS_E1_xyz99__DQ0" }
}
```

Instead of a static list, `answer_options` is a dict pointing to another question whose submitted values are used as the option list at runtime. Used to build dynamic dropdowns from previous answers (e.g. list of subsidiaries entered in a prior question group).

#### Required field

```json
{
  "required": true
}
```

Only present when `true` (omitted otherwise).

#### Conditional display

```json
{
  "condition": "ESRS_E1_abc12",
  "condition_criteria": "coal|gas",
  "nesting_level": 1
}
```

| Field | Description |
|---|---|
| `condition` | Logical expression of question IDs, e.g. `"Q1"`, `"(Q1|Q2)&Q3"`. Uses `&` (AND), `\|` (OR), `!` (NOT). |
| `condition_criteria` | Truth map defining which values of the referenced question(s) make this question visible. Can be a simple string (`"Yes\|Partly"`), a compound value using `&`/`\|`, or a JSON dict mapping multiple question IDs to their criteria: `{"Q1": "Yes", "Q2": "High\|Medium"}`. |
| `nesting_level` | Integer indicating visual indent depth (0 = top-level, 1+ = nested). `null` means auto-calculated from condition structure. |

---

### Dynamic / repeatable group questions

Questions inside a repeatable `QuestionGroup` carry additional fields that control the group's rendering:

```json
{
  "id": "ESRS_E1_abc12__DQ0",
  "dynamic_group_id": "subsidiaries_group",
  "dynamic_header": "Subsidiary",
  "dynamic_max_items": 10,
  "dynamic_tag_expression": "{sub_name__DQ0} ({sub_country__DQ0})",
  "dynamic_condition": "has_subsidiaries",
  "dynamic_condition_criteria": "Yes"
}
```

| Field | Description |
|---|---|
| `id` (with `__DQ0`) | The `__DQ0` suffix marks a **dynamic question template**. At runtime, the frontend replicates this template for each group instance, replacing `DQ0` with `DQ1`, `DQ2`, etc. |
| `dynamic_group_id` | Identifier for the repeatable group all questions in the same group share. |
| `dynamic_header` | Label shown as the accordion/section header for each group instance. |
| `dynamic_max_items` | Maximum number of repeatable instances. `null` = unlimited. `1` = fixed/non-repeatable. |
| `dynamic_tag_expression` | Template string for accordion item labels. Uses `{question_id}` placeholders resolved from sibling question values. E.g. `"{first_name__DQ0} {last_name__DQ0}"`. If absent, the first question's value is used. |
| `dynamic_condition` | Condition expression (same syntax as `condition`) that controls **when the entire group is shown**. |
| `dynamic_condition_criteria` | Truth map for `dynamic_condition`. |

In the **proxy definition**, all dynamic question templates use `__DQ0`. In **saved ticket data**, each actual instance uses `__DQ1`, `__DQ2`, etc.

---

## 2. Answer Types

Defined in `utils/data_types.py`:

| Value | Label | Notes |
|---|---|---|
| `OPEN` | Open Text | Single-line text input |
| `OPEN_LARGE` | Open Large Text | Multi-line / textarea |
| `INTEGER` | Integer | Numeric integer input |
| `NUMERIC` | Numeric | Decimal number input |
| `RADIO` | Radio Buttons | Single select from `answer_options` list |
| `DROPDOWN` | Dropdown | Single select from `answer_options` list |
| `CHECKBOX` | Checkboxes | Multi-select from `answer_options` list |
| `RICH_TEXT` | Rich Text | TipTap WYSIWYG editor |
| `DATE` | Date | Date picker |
| `TIME` | Time | Time picker |
| `SCALE` | Scale/Rating | Numeric scale between `scale_min` and `scale_max` |
| `EMAIL` | Email | Validated email input |
| `URL` | URL | Validated URL input |
| `LABEL` | Label | Display-only text, no input (max 50 chars) |
| `RADIO_REF` | Radio (Reference) | Single select; options populated at runtime from another question's answers |
| `DROPDOWN_REF` | Dropdown (Reference) | Same as above, dropdown variant |
| `CHECKBOX_REF` | Checkboxes (Reference) | Same as above, multi-select variant |

**Legacy format:** Older imported JSON may encode choice questions as `"answer_type": "Yes/No/Partly"` (slash-delimited). The import command converts these to `DROPDOWN` with equivalent `answer_options`.

---

## 3. Ticket Response JSON (`Ticket.survey_data`)

When a respondent fills in a survey, the ticket stores a **flat list of answered question objects**. Each object is a copy of the corresponding question from `survey_data_proxy`, enriched with the respondent's input:

```json
[
  {
    "id": "ESRS_E1_abc12",
    "name": "What is your primary energy source?",
    "answer_type": "RADIO",
    "value": "solar",
    "answer_options": [
      { "value": "coal",  "label": "Coal" },
      { "value": "gas",   "label": "Natural Gas" },
      { "value": "solar", "label": "Solar" }
    ],
    "help_text": "Select the most relevant option.",
    "related_ids": [],
    "completed": true
  }
]
```

Key differences from the proxy definition:

| Field | Description |
|---|---|
| `value` | The respondent's answer. For `CHECKBOX`, a list of selected values. For `_REF` types, the chosen value(s) from the referenced question's instances. |
| `completed` | `true` when the question has been answered and saved. Used for progress tracking. |
| `comments` | Optional string; present if the respondent added a comment to the question. |
| `tag` | Optional; present on dynamic group questions (`__DQn`). Stores the resolved label for this group instance (used in accordion headers). |

### Dynamic group instances in ticket data

For repeatable groups, the ticket stores one entry **per instance** per question, using incrementing suffixes:

```json
[
  { "id": "sub_name__DQ1", "dynamic_group_id": "subsidiaries_group", "value": "Acme Ltd", "completed": true },
  { "id": "sub_country__DQ1", "dynamic_group_id": "subsidiaries_group", "value": "SE", "completed": true },
  { "id": "sub_name__DQ2", "dynamic_group_id": "subsidiaries_group", "value": "Beta GmbH", "completed": true },
  { "id": "sub_country__DQ2", "dynamic_group_id": "subsidiaries_group", "value": "DE", "completed": true }
]
```

The `__DQ0` template never appears in saved ticket data — only `__DQ1`, `__DQ2`, etc.

---

## 4. Runtime Merge (`get_latest_survey_data_version`)

Defined in `home/utils.py`, this function produces the final merged list used for both display and reporting.

```
survey_data_proxy  (definition)   ──┐
                                     ├──► get_latest_survey_data_version() ──► merged survey_data
ticket.survey_data (user answers)  ──┘
```

The merge performs the following, per question item:

1. **Base**: Start from the ticket's saved item.
2. **Proxy lookup**: Normalise the item's ID (strip `__DQn` → `__DQ0`) and look up the matching proxy item.
3. **Cosmetic patching** (same version only, incomplete items):
   - Overwrite `name` and `help_text` from the proxy (picks up text corrections applied via a patch).
   - Update `label` values in `answer_options` from the proxy (preserving the stored `value`).
   - Sync structural/behavioural metadata from the proxy: `dynamic_tag_expression`, `dynamic_max_items`, `dynamic_header`, `dynamic_group_id`, `dynamic_condition`, `dynamic_condition_criteria`, `nesting_level`, `required`, `related_ids`, `scale_min`, `scale_max`, `scale_step`.
4. **Version mismatch**: If `survey_data_proxy_version` differs between the ticket and the survey, proxy patching is skipped entirely and the ticket data is returned as-is (backward compatibility).
5. **Condition enforcement** (reporting mode): If `enforce_conditions=True`, questions whose `condition` is not satisfied (given the current answer values) have their `value` replaced with `"_n/a_"`.
6. **Progress counting**: Counts unique main question IDs with `completed=true` for the completion percentage.

The `survey_data_proxy_lookup` passed to this function is a dict keyed by `question_id` (`__DQ0` form), built from `Survey.survey_data_proxy`.

---

## 5. How Online Surveys Are Rendered

### Server-side (view layer)

Two view classes handle survey rendering:

- **`SurveyView`** (`home/views/survey.py`) — authenticated users (board members, respondents).
- **`GuestSurveyView`** (`home/views/survey.py`) — external respondents via a JWT link (no login required). The JWT encodes `ticket_id` and `org_id`.

Both views:
1. Load `Survey.survey_data_proxy` and build `survey_data_proxy_lookup` (dict keyed by question ID).
2. Load `Ticket.survey_data` (current user answers).
3. Call `get_latest_survey_data_version()` to produce the merged survey data.
4. Pass the merged list to the template as `survey_data`.

Auto-save (PATCH) and submit (POST) are handled by the same views. On each save, the full merged `survey_data` list is written back to `Ticket.survey_data`.

### Client-side (JavaScript)

The survey UI is built in `SurveyCarouselManager` (extends `CarouselManager`), bundled via Webpack into `webpack_bundle.js`.

**Rendering flow:**

1. The Django template renders `survey_data` questions as hidden `<div class="question-item">` elements inside the HTML.
2. `SurveyCarouselManager.renderQuestions()` finds all `.question-item` elements, strips admin-only UI (mark-complete buttons, action bars), and calls `buildSlide()` for each.
3. `buildSlide()` wraps each question in a Bootstrap carousel slide `<div class="carousel-item">` with Previous / Continue navigation buttons.
4. The last **visible** slide's Continue button is replaced by a green **Submit survey** button by `updateSubmitButton()`, which re-runs whenever conditional visibility changes.
5. The carousel shows one question at a time. Navigation (Previous/Continue) moves between slides, updating the progress bar via `updateProgressBar()`.

**Conditional visibility:**

Condition logic is evaluated client-side using `condition_is_active()`. When an answer changes:
- Questions with a failing `condition` are hidden (and their slide is skipped in the carousel).
- `updateSubmitButton()` is called to ensure Submit always lands on the correct final visible question.
- Dynamic group questions with a failing `dynamic_condition` are hidden/collapsed.

**Dynamic (repeatable) groups:**

The frontend handles `__DQ0` templates by:
- Rendering the `__DQ0` question as an "Add item" template.
- On each "Add" click, cloning the template elements, incrementing the DQ counter (`DQ1`, `DQ2`, ...), and inserting the new instance.
- The `dynamic_tag_expression` is resolved at runtime to build accordion labels from sibling question values.
- `dynamic_max_items` caps the maximum number of instances.

**Auto-save:**

On Continue (slide change), `SurveyCarouselManager` triggers an auto-save via a `fetch` POST to the survey endpoint, sending the current `survey_data` as JSON. The backend validates and persists the data to `Ticket.survey_data`.

**Submission:**

On Submit, the same fetch is called with a `submit=true` flag. The server:
1. Validates the full `survey_data`.
2. Sets `Ticket.status = "submitted"`.
3. Creates a `TicketHistory` record.
4. Sends a notification email to the inviter (if applicable).
5. The client navigates to the completion slide.

---

## 6. Survey Version Control & Data Integrity

### Version tracking

`Survey.survey_data_proxy_version` tracks which `SurveyVersion.version_number` the current proxy was taken from. `Ticket.survey_data_proxy_version` records the version active when the ticket was last saved.

If these two values differ, `get_latest_survey_data_version()` detects a **version mismatch** and skips cosmetic patching, preserving the ticket's original structure.

### Patch vs new version

| Change type | Mechanism | Effect on tickets |
|---|---|---|
| Cosmetic (text, labels) | **Patch** — `SurveyVersion.apply_patch()` | Ticket's `name`/`help_text`/labels updated on next load via cosmetic patching in merge |
| Structural (add/remove questions, change types, reorder) | **New version** — `SurveyVersion.publish()` on a new draft | Existing tickets retain old structure; new tickets use new version |

### Version migration

When upgrading an existing ticket to a new survey version, `migrate_survey_data_values()` (`home/utils.py`) handles the transition:
- Static questions are matched by ID and their values are type-validated and coerced.
- Dynamic group instances are matched by group + DQ suffix and replicated into the new structure.
- Removed questions are logged but their values are discarded.
- New questions are initialised with empty values.

---

## 7. Complete Example

A minimal two-question survey with one conditional and one repeatable group:

```json
[
  {
    "id": "SURVEY_q1abc",
    "name": "Does your company have subsidiaries?",
    "value": "",
    "answer_type": "RADIO",
    "answer_options": [
      { "value": "Yes", "label": "Yes" },
      { "value": "No",  "label": "No" }
    ],
    "help_text": "",
    "related_ids": [],
    "required": true
  },
  {
    "id": "SURVEY_q2def__DQ0",
    "name": "Subsidiary name",
    "value": "",
    "answer_type": "OPEN",
    "help_text": "Legal name of the subsidiary.",
    "related_ids": [],
    "dynamic_group_id": "subsidiaries_group",
    "dynamic_header": "Subsidiary",
    "dynamic_tag_expression": "{SURVEY_q2def__DQ0}",
    "dynamic_condition": "SURVEY_q1abc",
    "dynamic_condition_criteria": "Yes",
    "required": true
  },
  {
    "id": "SURVEY_q3ghi__DQ0",
    "name": "Subsidiary country",
    "value": "",
    "answer_type": "DROPDOWN",
    "answer_options": [
      { "value": "SE", "label": "Sweden" },
      { "value": "DE", "label": "Germany" }
    ],
    "help_text": "",
    "related_ids": [],
    "dynamic_group_id": "subsidiaries_group",
    "dynamic_header": "Subsidiary",
    "dynamic_tag_expression": "{SURVEY_q2def__DQ0}",
    "dynamic_condition": "SURVEY_q1abc",
    "dynamic_condition_criteria": "Yes",
    "required": true
  }
]
```

After a respondent adds two subsidiaries and answers the first question "Yes", `Ticket.survey_data` would contain:

```json
[
  {
    "id": "SURVEY_q1abc",
    "name": "Does your company have subsidiaries?",
    "answer_type": "RADIO",
    "value": "Yes",
    "answer_options": [{ "value": "Yes", "label": "Yes" }, { "value": "No", "label": "No" }],
    "help_text": "",
    "related_ids": [],
    "required": true,
    "completed": true
  },
  {
    "id": "SURVEY_q2def__DQ1",
    "answer_type": "OPEN",
    "dynamic_group_id": "subsidiaries_group",
    "value": "Acme Ltd",
    "tag": "Acme Ltd",
    "completed": true
  },
  {
    "id": "SURVEY_q3ghi__DQ1",
    "answer_type": "DROPDOWN",
    "dynamic_group_id": "subsidiaries_group",
    "value": "SE",
    "completed": true
  },
  {
    "id": "SURVEY_q2def__DQ2",
    "answer_type": "OPEN",
    "dynamic_group_id": "subsidiaries_group",
    "value": "Beta GmbH",
    "tag": "Beta GmbH",
    "completed": true
  },
  {
    "id": "SURVEY_q3ghi__DQ2",
    "answer_type": "DROPDOWN",
    "dynamic_group_id": "subsidiaries_group",
    "value": "DE",
    "completed": true
  }
]
```

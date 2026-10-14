# Test Scenario CSV Converter Agent

## Overview
This agent converts test scenarios from Excel/table format into a standardized CSV template for Jira bulk import. It handles single and multi-scenario test cases with flexible Parent ID assignment.

## Usage

### Basic Command
```bash
python test_scenario_converter.py --input <file_or_text> --output <output_file>
```

### Required Parameters
- `--input` — Path to Excel file, CSV file, or pasted table/text
- `--output` — Output CSV filename (default: test_scenarios.csv)

### Optional Parameters
- `--parent-ids` — Parent ID assignment (see below)
- `--assignee` — Jira assignee (default: prompts for input)
- `--test-type` — Type of testing: Manual, Automated, etc. (default: prompts)
- `--fix-version` — Fix Version (default: prompts if not in source)
- `--no-fix-version` — Skip Fix Version if not provided

## Parent ID Assignment Options

The agent accepts flexible Parent ID input:

### Option 1: Use Same Parent ID for All Test Cases
```bash
python test_scenario_converter.py --input data.xlsx --parent-ids "ADXP-5439 for all"
```

### Option 2: Use Different IDs by Range
```bash
python test_scenario_converter.py --input data.xlsx --parent-ids "ADXP-5439 for 001-005, ADXP-5440 for 006-010"
```

### Option 3: Per Test Case (in Source Data)
Include a "Parent ID" column in your source Excel/CSV. The agent will extract and use it.

### Option 4: Interactive (Agent Asks)
```bash
python test_scenario_converter.py --input data.xlsx
```
The agent will ask for Parent ID assignment for each test case.

## Input Formats Supported

### 1. Excel File (.xlsx)
```bash
python test_scenario_converter.py --input scenarios.xlsx --parent-ids "ADXP-5439 for all"
```

### 2. CSV File
```bash
python test_scenario_converter.py --input scenarios.csv --parent-ids "ADXP-5439 for all"
```

### 3. Pasted Table/Text
```bash
python test_scenario_converter.py --input "paste_table_or_text_here" --parent-ids "ADXP-5439 for all"
```

## Input Data Requirements

Your source data should include (at minimum):
- **Test Case ID** — Unique identifier (001, 002, etc.)
- **Scenario Description** — What the test validates
- **Test Steps** — Numbered steps (1. Step, 2. Step, etc.)
- **Expected Result** — What should happen
- **Pre-Condition** (optional) — Setup required before test
- **Test Data** (optional) — User accounts, test values, etc.
- **User Story ID** (optional) — Associated story/epic
- **Parent ID** (optional) — If in source; otherwise provide via --parent-ids

### Multiple Scenarios per Test Case
If a test case has multiple scenarios, include them with clear headings:
```
Test Case 001
SCENARIO 1 - Account Login
[steps and expected results for scenario 1]

SCENARIO 2 - Session Management
[steps and expected results for scenario 2]
```

The agent will combine them into one CSV row with `SCENARIO 1 -` and `SCENARIO 2 -` headings.

## Output

The agent generates a **UTF-8 encoded CSV** with these 21 columns:

```
Parent ID | IssueType | Assignee | Fix Version | User Story ID | Scenario Type | 
Region Tested | TestType | Test Scenario/Test Case ID | Summary | Acceptance Criteria | 
Pre-Condition | Test Scenario Description | Test Steps | Expected Result | Test Data | 
QA Planned Start Date | QA Planned End Date | QA Actual Start Date | QA Actual End Date | Priority
```

### Default Values Applied Automatically
- **IssueType**: Test Case
- **Scenario Type**: Functional
- **Region Tested**: QA
- **Priority**: Medium
- **QA Date Fields**: Blank (fill later in Jira)

## Examples

### Example 1: Simple Conversion with Same Parent ID
```bash
python test_scenario_converter.py \
  --input scenarios.xlsx \
  --parent-ids "ADXP-5439 for all" \
  --assignee john_smith \
  --test-type Manual \
  --fix-version "Dashboard v2.1"
```

### Example 2: Multiple Parent IDs by Range
```bash
python test_scenario_converter.py \
  --input test_cases.csv \
  --parent-ids "ADXP-5439 for 001-005, ADXP-5440 for 006-010" \
  --assignee team_lead \
  --test-type Automated
```

### Example 3: Interactive (Agent Asks for Input)
```bash
python test_scenario_converter.py --input data.xlsx
```
Agent will prompt for:
- Parent ID assignment
- Assignee
- TestType
- Fix Version

## Integration with Claude Code (VSCode)

1. **Save the script** in your project:
   ```
   /your-project/test_scenario_converter.py
   ```

2. **Run via Claude Code**:
   - In VSCode, open Claude Code chat
   - Type: `Run test_scenario_converter.py --input scenarios.xlsx --parent-ids "ADXP-5439 for all"`
   - Claude Code will execute it and show results

3. **Or run manually** from terminal:
   ```bash
   cd /your-project
   python test_scenario_converter.py --input scenarios.xlsx --parent-ids "ADXP-5439 for all"
   ```

## Integration with Copilot (Custom Agent)

If using Copilot with custom agent support:

1. Place `agent.md` (this file) in your project root
2. Reference in Copilot: `@test-scenario-converter` or `@agent` depending on your setup
3. Copilot will read this file and use it to understand the agent's capabilities
4. Ask Copilot to run the conversion:
   - "Convert my test scenarios to Jira CSV using the test scenario converter"
   - "Run the converter on scenarios.xlsx with parent ID ADXP-5439"

Copilot will then call the Python script with appropriate parameters.

## Error Handling

The agent will ask for clarification if:
- **Source data is messy** — Inconsistent formatting, missing required fields
- **Parent ID is ambiguous** — Multiple IDs needed but not specified
- **Fix Version is missing** — Asks if you want to skip or provide one
- **Test Case ID is unclear** — Asks for confirmation

The agent will **NOT guess or assume** values.

## Output Verification

Before importing to Jira, verify:
- ✓ Each test case has correct Parent ID
- ✓ Assignee is populated
- ✓ TestType matches your input
- ✓ Fix Version is correct
- ✓ Multi-scenario test cases have clear `SCENARIO 1 -` and `SCENARIO 2 -` headings
- ✓ All test data is preserved exactly as written
- ✓ CSV is UTF-8 encoded

## Jira Import

Once the CSV is generated:

1. Open Jira → Issues → Import
2. Select CAV (Create/Associate Values) feature
3. Upload the test_scenarios.csv
4. Map columns to Jira fields using your custom mapping rule
5. Review and import

The CSV is ready for direct import — no post-processing needed.

## Troubleshooting

### "ModuleNotFoundError: No module named 'openpyxl'"
Install required packages:
```bash
pip install openpyxl pandas
```

### "Parent ID not recognized"
Verify format:
- `"ADXP-5439 for all"` ✓
- `"ADXP-5439 for 001-005, ADXP-5440 for 006-010"` ✓
- `"ADXP-5439"` (for single test case) ✓

### CSV has extra empty rows
This is normal if your source data has blank rows. Jira will handle them during import.

### Multi-scenario separation unclear
Check that scenario headings use exact format:
```
SCENARIO 1 - [Name]
[content]

SCENARIO 2 - [Name]
[content]
```

## Support

If the agent fails to parse your input:
1. Provide sample data (sanitized) to debug
2. Specify exact Parent ID assignment logic
3. Check source data formatting (Excel columns, table structure, etc.)

The agent is designed to preserve data exactly as written — it won't interpret or reword scenarios.

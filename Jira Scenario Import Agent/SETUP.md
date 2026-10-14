# Test Scenario Converter Agent - Setup Guide

This guide shows you how to set up and use the Test Scenario CSV Converter Agent with Claude Code in VSCode or with Copilot.

## Quick Start

### Prerequisites
- Python 3.7+
- Claude Code extension in VSCode (or Copilot)
- Optional: `openpyxl` or `pandas` for Excel support

### Installation

1. **Download the files**:
   - `test_scenario_converter.py` — Main converter script
   - `agent.md` — Agent documentation
   
2. **Place in your project**:
   ```
   your-project/
   ├── test_scenario_converter.py
   ├── agent.md
   └── test_data/
       └── scenarios.xlsx (your test data)
   ```

3. **Install dependencies** (for Excel support):
   ```bash
   pip install openpyxl pandas
   ```
   
   If you only need CSV/text support, you can skip this.

---

## Integration with Claude Code (VSCode)

Claude Code lets you run Python scripts directly from the chat within VSCode.

### Setup Steps

1. **Open VSCode** with your project
2. **Install Claude Code extension** (if not already installed)
3. **Open Claude Code chat** (Ctrl+K in VSCode)
4. **Ask Claude Code to run the script**:

### Example: Run the Converter

**In Claude Code chat, type:**

```
Run the test scenario converter with scenarios.xlsx, use ADXP-5439 for all parent IDs, assignee is john_smith
```

**Or run directly:**

```bash
python test_scenario_converter.py --input test_data/scenarios.xlsx --parent-ids "ADXP-5439 for all" --assignee john_smith --test-type Manual --fix-version "Dashboard v2.1"
```

### Common Claude Code Commands

**Interactive mode** (Claude Code asks for input):
```bash
python test_scenario_converter.py --input test_data/scenarios.xlsx
```

**Batch mode** (all params specified):
```bash
python test_scenario_converter.py \
  --input test_data/scenarios.xlsx \
  --parent-ids "ADXP-5439 for 001-005, ADXP-5440 for 006-010" \
  --assignee john_smith \
  --test-type Manual \
  --fix-version "Dashboard v2.1" \
  --output test_scenarios.csv
```

**From pasted data**:
```bash
python test_scenario_converter.py \
  --input "Test Case 001\nStep 1: Login\nStep 2: Verify" \
  --parent-ids "ADXP-5439 for all"
```

---

## Integration with Copilot (Custom Agent)

If using Copilot with custom agent support, you can reference this agent.

### Setup Steps

1. **Place agent.md in your workspace**:
   ```
   your-project/
   ├── agent.md
   ├── test_scenario_converter.py
   └── test_data/
   ```

2. **Reference in Copilot**:
   - Type: `@test-scenario-converter` or reference the agent by name
   - Copilot reads `agent.md` to understand capabilities
   - Ask Copilot to run the converter

### Example Copilot Prompts

```
@agent Convert my test scenarios to Jira CSV using the test scenario converter. 
File is at test_data/scenarios.xlsx, use ADXP-5439 for all test cases.
```

Or:

```
@test-scenario-converter
I have test cases in scenarios.csv. Convert them to Jira format with:
- Parent ID: ADXP-5439 for 001-005, ADXP-5440 for 006-010
- Assignee: qa_team
- Test Type: Manual
```

---

## Usage Examples

### Example 1: Simple Conversion (Same Parent ID)

**Command:**
```bash
python test_scenario_converter.py \
  --input scenarios.xlsx \
  --parent-ids "ADXP-5439 for all" \
  --assignee john_smith \
  --test-type Manual \
  --fix-version "v2.1"
```

**Input**: `scenarios.xlsx` with columns:
| Test Case ID | Scenario Description | Test Steps | Expected Result |
|---|---|---|---|
| 001 | Account Login | 1. Navigate... | User logged in... |
| 002 | Dashboard | 1. Click... | Dashboard loaded... |

**Output**: `test_scenarios.csv` with 21 columns, ready for Jira import

### Example 2: Multiple Parent IDs by Range

**Command:**
```bash
python test_scenario_converter.py \
  --input test_cases.csv \
  --parent-ids "ADXP-5439 for 001-005, ADXP-5440 for 006-010, ADXP-5441 for 011+" \
  --assignee qa_team \
  --test-type Automated
```

Maps:
- Test cases 001-005 → Parent ID ADXP-5439
- Test cases 006-010 → Parent ID ADXP-5440
- Test case 011+ → Parent ID ADXP-5441

### Example 3: Interactive Mode (Agent Asks)

**Command:**
```bash
python test_scenario_converter.py --input scenarios.xlsx
```

**Agent will ask for:**
1. Parent ID assignment
2. Assignee
3. Test Type
4. Fix Version (if not in source)

### Example 4: Pasted Table/Text

**Command:**
```bash
python test_scenario_converter.py --input "Test Case 001\nStep 1: Login\nExpected: User logged in" --parent-ids "ADXP-5439 for all" --assignee john
```

---

## Input Format Guide

### Supported Input Formats

#### 1. Excel File (.xlsx)

**Required columns:**
- `Test Case ID` (e.g., 001, TC-001, TEST-001)
- `Test Scenario Description` or `Scenario`
- `Test Steps` (can be multi-line)
- `Expected Result`

**Optional columns:**
- `Pre-Condition`
- `Test Data`
- `Acceptance Criteria`
- `User Story ID`
- `Parent ID` (if per-test-case)
- `Fix Version`

#### 2. CSV File

Same column requirements as Excel.

#### 3. Pasted Text/Table

Format examples:
```
Test Case 001 - Account Login
Pre-Condition: User not logged in
Test Steps:
1. Navigate to login page
2. Enter username
3. Enter password
4. Click Login
Expected Result:
- Login successful
- User redirected to dashboard
Test Data: username=john, password=test123
```

Or simple table:
```
| Test Case | Steps | Expected |
|-----------|-------|----------|
| 001 | Navigate to login | Page loads |
| 002 | Click login | User logged in |
```

### Multiple Scenarios per Test Case

If a test case has multiple scenarios, format as:

```
Test Case 001 - Dashboard Navigation

SCENARIO 1 - Single DAF
Pre-Condition: User has single DAF
Test Steps:
1. Login as single DAF user
2. Navigate to dashboard
Expected Result: Dashboard loads directly

SCENARIO 2 - Multiple DAF
Pre-Condition: User has multiple DAFs
Test Steps:
1. Login as multi-DAF user
2. Navigate to account module
Expected Result: DAF selection dialog appears
```

The agent will combine both scenarios into **one CSV row** with scenario headings preserved.

---

## Output Verification

Once the CSV is generated, verify before import:

### Check the CSV

```bash
# View first few rows
head -5 test_scenarios.csv

# Check row count
wc -l test_scenarios.csv

# View specific column (e.g., Parent ID)
cut -d',' -f1 test_scenarios.csv | head -10
```

### CSV Structure

The output will have these 21 columns (in order):
1. Parent ID
2. IssueType (always "Test Case")
3. Assignee
4. Fix Version
5. User Story ID
6. Scenario Type (always "Functional")
7. Region Tested (always "QA")
8. TestType (Manual/Automated/etc.)
9. Test Scenario/Test Case ID
10. Summary
11. Acceptance Criteria
12. Pre-Condition
13. Test Scenario Description
14. Test Steps
15. Expected Result
16. Test Data
17. QA Planned Start Date (blank)
18. QA Planned End Date (blank)
19. QA Actual Start Date (blank)
20. QA Actual End Date (blank)
21. Priority (always "Medium")

---

## Jira Import

Once verified, import to Jira:

1. **Open Jira** → Navigate to Issues
2. **Click Import** (CAV feature)
3. **Upload CSV**: Select `test_scenarios.csv`
4. **Map columns**: Use your custom mapping rule
5. **Review**: Check test cases look correct
6. **Import**: Click Import button

The CSV is **UTF-8 encoded and ready** for direct import.

---

## Troubleshooting

### "ModuleNotFoundError: No module named 'openpyxl'"

**Solution**: Install Excel support
```bash
pip install openpyxl pandas
```

If you only need CSV/text support, ignore this error.

### "Error: No test cases found in input"

**Likely cause**: Input file is empty or format not recognized

**Solutions**:
- Check file path is correct
- Ensure file has data
- Try copy/pasting data as text instead

### "Parent ID not recognized"

**Likely cause**: Parent ID format incorrect

**Valid formats**:
- ✓ `"ADXP-5439 for all"`
- ✓ `"ADXP-5439 for 001-005, ADXP-5440 for 006-010"`
- ✗ `"ADXP-5439, ADXP-5440"` (wrong format)
- ✗ `"ADXP-5439"` (missing "for all" or range)

### CSV has extra empty rows

**Cause**: Source data has blank rows

**Fix**: This is normal. Jira will ignore empty rows during import.

### Multi-scenario headings not showing

**Likely cause**: Scenario format not recognized

**Solution**: Use exact format:
```
SCENARIO 1 - [Name]
[content]

SCENARIO 2 - [Name]
[content]
```

---

## Advanced Usage

### Using Environment Variables

Set defaults to avoid typing repeatedly:

```bash
export JIRA_ASSIGNEE="john_smith"
export JIRA_TEST_TYPE="Manual"
export JIRA_FIX_VERSION="Dashboard v2.1"

python test_scenario_converter.py \
  --input scenarios.xlsx \
  --parent-ids "ADXP-5439 for all" \
  --assignee $JIRA_ASSIGNEE \
  --test-type $JIRA_TEST_TYPE \
  --fix-version $JIRA_FIX_VERSION
```

### Batch Processing

Convert multiple files:

```bash
for file in test_data/*.xlsx; do
  python test_scenario_converter.py \
    --input "$file" \
    --parent-ids "ADXP-5439 for all" \
    --assignee john_smith \
    --output "output_${file%.xlsx}.csv"
done
```

### Integrate with Git

Add to `.gitignore`:
```
test_scenarios.csv
test_data/*.xlsx
```

Keep `agent.md` and `test_scenario_converter.py` in version control.

---

## Support & Help

**If something doesn't work:**

1. Check the `agent.md` file for full documentation
2. Verify input format matches examples above
3. Try interactive mode to debug:
   ```bash
   python test_scenario_converter.py --input your_file
   ```
4. Check Parent ID assignment format

**Common issues:**
- Messy input → Agent will ask for clarification
- Missing Parent ID → Agent will ask
- Missing Fix Version → Agent will ask (can bypass with `--no-fix-version`)
- Inconsistent test case IDs → Agent normalizes them automatically

---

## What's Next?

Once you have the CSV:
1. ✓ Download/verify the CSV file
2. ✓ Import to Jira via Issues → Import (CAV feature)
3. ✓ Use custom mapping rule to map columns to Jira fields
4. ✓ Review test cases in Jira
5. ✓ Start testing!

Enjoy! 🚀

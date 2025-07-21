# CertSuite Claim Spreadsheet Generator

A Python tool that converts CertSuite claim.json files into well-formatted Excel spreadsheets for easy analysis and reporting.

## Overview

This script processes CertSuite test results from JSON claim files and generates comprehensive Excel reports with:
- Test results organized by status (Failed, Skipped, Passed)
- Summary statistics and test counts
- Version information for all components
- Color-coded formatting for easy analysis
- Detailed test output and remediation information

## Features

- **Automated Test Analysis**: Automatically categorizes and sorts tests by their status
- **Rich Excel Formatting**: Color-coded cells, proper column widths, and professional styling
- **Summary Dashboard**: Overview of test counts and success rates
- **Version Tracking**: Component version information from the claim file
- **DCI Integration**: Can automatically download claim files from DCI (Distributed CI) servers
- **Offline Mode**: Works with local claim.json files without DCI connectivity
- **Error Handling**: Robust error handling for malformed JSON and missing data

## Prerequisites

### Required Python Packages

```bash
pip install openpyxl
```

*Note: `argparse` is included in Python's standard library, so no separate installation is needed.*

### Optional: DCI Integration

For automatic claim file downloading from DCI servers:
- `dcictl` command-line tool installed and configured
- DCI credentials configured via either:
  - Environment variables (exported in your shell)
  - `dcirc.sh` file with export statements (automatically read by the script)

## Installation

1. Clone or download the script:
```bash
git clone <repository-url>
cd certsuite-claim-spreadsheet
```

2. Install required Python packages:
```bash
pip install -r requirements.txt
```

3. (Optional) Set up DCI credentials using either method:

   **Method 1: Environment variables (temporary)**
   ```bash
   export DCI_CLIENT_ID="your-client-id"
   export DCI_API_SECRET="your-api-secret"
   export DCI_CS_URL="https://api.distributed-ci.io"
   ```

   **Method 2: Create `dcirc.sh` file (recommended)**
   ```bash
   # Option A: Copy and edit the sample file
   cp dcirc.sh.sample dcirc.sh
   # Edit dcirc.sh with your actual credentials
   
   # Option B: Create from scratch
   cat > dcirc.sh << EOF
   export DCI_CLIENT_ID="your-client-id"
   export DCI_API_SECRET="your-api-secret"
   export DCI_CS_URL="https://api.distributed-ci.io"
   EOF
   ```
   
   The script will automatically read from `dcirc.sh` if environment variables aren't set.

## Usage

### Basic Usage (Offline Mode)

Process a local claim.json file:

```bash
python certsuite_claim_spreadsheet.py -i claim.json -o test_results.xlsx
```

### DCI Integration Mode

Download and process claim file from DCI server:

```bash
python certsuite_claim_spreadsheet.py -i claim.json -j <job-id> -o test_results.xlsx
```

### Command Line Arguments

| Argument | Short | Required | Description |
|----------|-------|----------|-------------|
| `--input-file` | `-i` | Yes | Path to the claim.json file (will be created if using DCI mode) |
| `--job-id` | `-j` | No | DCI job ID for automatic download. If not specified, uses offline mode |
| `--output-file` | `-o` | Yes | Path for the output Excel file (e.g., `results.xlsx`) |

### Examples

**Example 1: Offline Mode**
```bash
python certsuite_claim_spreadsheet.py -i /path/to/claim.json -o report.xlsx
```

**Example 2: DCI Mode**
```bash
python certsuite_claim_spreadsheet.py -i claim.json -j 12345 -o dci_report.xlsx
```

**Example 3: Custom Output Location**
```bash
python certsuite_claim_spreadsheet.py -i claim.json -o /reports/monthly_certsuite_$(date +%Y-%m).xlsx
```

## Input Format

The script expects a CertSuite claim.json file with the following structure:

```json
{
  "claim": {
    "results": {
      "test-id-1": {
        "testID": {"id": "test-name"},
        "state": "passed|failed|skipped|error",
        "catalogInfo": {
          "description": "Test description",
          "exceptionProcess": "Exception process",
          "remediation": "Remediation steps",
          "bestPracticeReference": "Best practice link"
        },
        "capturedTestOutput": "Test output logs",
        "categoryClassification": {
          "Extended": "true",
          "Telco": "false"
        }
      }
    },
    "versions": {
      "k8s": "v1.28.0",
      "ocClient": "4.14.0",
      "ocp": "4.14.0",
      "certSuite": "v5.0.0",
      "claimFormat": "0.4.0",
      "certSuiteGitCommit": "abc123"
    }
  }
}
```

## Output Format

The generated Excel file contains:

### Summary Section
- Total test count
- Failed test count (highlighted in red)
- Error test count (highlighted in dark red)
- Skipped test count (highlighted in orange)
- Passed test count (highlighted in green)
- DCI Job ID link (if applicable)

### Test Suite Summary (New Worksheet)
A separate "Suite Summary" worksheet with:
- **SUITE**: Test suite name (access-control, networking, etc.)
- **PASSED**: Number of passed tests per suite
- **FAILED**: Number of failed tests per suite  
- **ERROR**: Number of error tests per suite
- **SKIPPED**: Number of skipped tests per suite
- **TOTAL**: Total tests per suite

**Category Classification Summary:**
Breakdown of tests by certification category:
- **Extended specific tests only**: Tests specific to Extended certification with Mandatory/Optional counts
- **Far-Edge specific tests only**: Tests specific to Far-Edge certification with Mandatory/Optional counts  
- **Non-Telco specific tests only**: Tests specific to Non-Telco certification with Mandatory/Optional counts
- **Telco specific tests only**: Tests specific to Telco certification with Mandatory/Optional counts

Example output:
```
-----------------------------------------------------------
| SUITE                       PASSED    FAILED    ERROR   SKIPPED | TOTAL |
-----------------------------------------------------------
| access-control                   9        11        2          7 |    29 |
| networking                       4         4        0          4 |    12 |
-----------------------------------------------------------

Category Classification Summary
-----------------------------------------------------------
| Category                    | Total | Mandatory | Optional |
-----------------------------------------------------------
| Extended specific tests only|   13  |    10     |    3     |
| Far-Edge specific tests only|    9  |     8     |    1     |
| Non-Telco specific tests only|  70  |    43     |   27     |
| Telco specific tests only   |   27  |    26     |    1     |
-----------------------------------------------------------
```

### Version Information
- Kubernetes version
- OpenShift Client version
- OpenShift Platform version
- CertSuite version
- Claim format version
- Git commit hash

### Test Results Table
- **Test ID**: Unique identifier for each test
- **Test Text**: Human-readable description
- **State**: Test status (Failed/Error/Skipped/Passed)
- **Capture Output**: Test execution logs (for failed/error/skipped tests)
- **Category Classification**: Test categories (Extended, Telco, etc.)
- **Exception Process**: Exception handling information
- **Remediation**: Steps to fix failed tests
- **Best Practice Link**: Reference documentation

## Color Coding

- **Red**: Failed tests and failed count
- **Dark Red**: Error tests and error count
- **Orange**: Skipped tests and skipped count
- **Green**: Passed tests and passed count
- **Blue**: Headers and summary labels
- **Yellow**: Version information section
- **Light Blue**: Special text highlighting for categories

## File Structure

```
certsuite-claim-spreadsheet/
├── certsuite_claim_spreadsheet.py  # Main script
├── dcirc.sh                        # DCI environment configuration (user-created)
├── dcirc.sh.sample                 # Sample DCI configuration file
├── README.md                       # This file
└── requirements.txt               # Python dependencies
```

**Generated Excel file structure:**
- **Sheet 1**: Detailed test results with summary and version info
- **Sheet 2**: Test suite summary with pass/fail/error/skip counts per suite

## Troubleshooting

### Common Issues

**1. "File not found" error**
```bash
Error: Input file not found: claim.json
```
- Ensure the claim.json file exists in the specified path
- Check file permissions

**2. "Invalid JSON" error**
```bash
Error: Invalid JSON in file claim.json
```
- Verify the JSON file is not corrupted
- Check for proper JSON formatting

**3. DCI connection issues**
```bash
dcictl command failed to download claim.json
```
- Verify `dcictl` is installed and configured
- Check DCI credentials are properly set (environment variables or `dcirc.sh`)
- Ensure network connectivity to DCI server
- The script will automatically try to read from `dcirc.sh` if environment variables aren't set

**4. DCI environment variable issues**
```bash
Missing environment variables: DCI_CLIENT_ID, DCI_API_SECRET
```
- The script will automatically try to read from `dcirc.sh` if environment variables aren't set
- Ensure your `dcirc.sh` file has the correct format with `export` statements
- Check that the `dcirc.sh` file is in the same directory as the script

**5. Missing test results**
```bash
Error: No test results found in claim data
```
- Verify the claim.json file contains the `claim.results` section
- Check if the file is a valid CertSuite claim file

### Debug Mode

For debugging, you can add print statements or use Python's logging module. The script already includes warning messages for non-critical errors.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the terms specified in the LICENSE file.

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Review the command-line help: `python certsuite_claim_spreadsheet.py --help`
3. Open an issue in the repository

## Changelog

### Latest Version
- **Added Category Classification Summary**: Suite Summary worksheet now includes breakdown by Extended/Far-Edge/Non-Telco/Telco with Mandatory/Optional counts
- **Enhanced mandatory/optional detection**: Improved logic using test tags, exception processes, and sophisticated pattern matching
- **Added Test Suite Summary**: New worksheet showing test counts by suite including error status
- **Added error status support**: Summary now includes error test count with dark red highlighting
- **Removed pandas dependency**: Eliminated unnecessary pandas usage, reducing dependencies and improving performance
- **Automatic dcirc.sh reading**: Script now automatically reads DCI credentials from `dcirc.sh` file
- **Improved DCI integration**: Better error handling and user feedback for DCI operations
- **Enhanced credential management**: Supports both environment variables and configuration file
- Refactored code into modular functions
- Added comprehensive error handling
- Improved type safety with type hints
- Enhanced Excel formatting and styling
- Added version information display
- Improved offline mode support

## Customizing Mandatory/Optional Classification

The script uses intelligent heuristics to determine if tests are mandatory or optional:

### **Detection Methods** (in order of priority):
1. **Test Tags**: Checks for 'mandatory', 'required', 'optional', 'informative' in test tags
2. **Exception Process**: Analyzes exception process text for keywords like "best practice", "recommendation"  
3. **Test ID Patterns**: Looks for explicit keywords in test names
4. **Test Description**: Analyzes test descriptions for mandatory/optional indicators
5. **Default Classification**: Most certification tests default to mandatory unless marked otherwise

### **Customization Options:**
You can customize the mandatory/optional logic by modifying the `is_test_mandatory()` function in the script to match your specific certification requirements.

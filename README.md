# 🤖 Chat Translator Test Automation

> **Automated end-to-end test suite for [PixelsSuite Chat Translator](https://www.pixelssuite.com/chat-translator)**
> — built with Python & Playwright, driven entirely by an Excel test case sheet.

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/Playwright-1.x-2EAD33?logo=playwright&logoColor=white)](https://playwright.dev/python/)
[![openpyxl](https://img.shields.io/badge/openpyxl-3.x-217346?logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 📖 Table of Contents

1. [Project Overview](#-project-overview)
2. [Features](#-features)
3. [Technologies Used](#-technologies-used)
4. [Folder Structure](#-folder-structure)
5. [Prerequisites](#-prerequisites)
6. [Installation](#-installation)
7. [Installing Playwright Browsers](#-installing-playwright-browsers)
8. [Running the Tests](#-running-the-tests)
9. [Command-Line Parameters](#-command-line-parameters)
10. [Excel Sheet Format](#-excel-sheet-format)
11. [Sample Workflow](#-sample-workflow)
12. [Expected Output](#-expected-output)
13. [Troubleshooting](#-troubleshooting)
14. [Future Improvements](#-future-improvements)
15. [Conclusion](#-conclusion)

---

## 🌐 Project Overview

This project is an **automated test automation framework** designed to validate the Singlish-to-Sinhala transliteration feature of the **PixelsSuite Chat Translator** web application.

Test cases are defined in a plain Excel workbook (`.xlsx`). The framework:

1. Reads each test row from the Excel sheet
2. Opens the Chat Translator in a Chromium browser (via Playwright)
3. Types the Singlish input into the text area
4. Clicks the **Transliterate** button and waits for the output
5. Compares the actual Sinhala output against the expected value
6. Writes `PASS`, `FAIL`, `COLLECTED`, or `UI Error` back into the Excel file

The result is a fully annotated Excel report produced **without any manual intervention**.

---

## ✨ Features

| Feature | Description |
|---|---|
| 📄 **Excel-driven tests** | All inputs and expected outputs are managed in a `.xlsx` file — no hard-coded test data |
| 🔍 **Smart column detection** | Automatically detects header rows and column names using fuzzy matching |
| 🔁 **Configurable retries** | Retries output reading if the translation is slow or hasn't updated yet |
| ⏱️ **Typing delay simulation** | Simulates realistic human typing speed to avoid triggering bot-detection |
| 🐌 **Slow-motion mode** | Slows every browser action for easy visual debugging |
| 💾 **Incremental auto-save** | Saves the workbook every *N* rows so progress is never lost on crash |
| 🖥️ **Headless / headed mode** | Run silently in CI or watch every action live in the browser |
| 🚪 **Keep-open mode** | Leaves the browser open after all tests for manual inspection |
| 🛡️ **Overlay dismissal** | Auto-closes cookie consent and GDPR overlays before interacting |
| 🔧 **Fully CLI-configurable** | Every behaviour is tunable via command-line arguments — no code changes needed |

---

## 🛠️ Technologies Used

| Technology | Version | Purpose |
|---|---|---|
| **Python** | 3.10+ | Core scripting language |
| **Playwright (sync)** | 1.x | Browser automation & UI interaction |
| **openpyxl** | 3.x | Reading and writing `.xlsx` Excel files |
| **argparse** | stdlib | Command-line argument parsing |
| **re / pathlib** | stdlib | Regex matching and path resolution |

---

## 📁 Folder Structure

```
test_automation/                        ← Project root
│
├── test_automation/                    ← Main package directory
│   ├── test_automation.py              ← Core automation script
│   └── Assignment 1 - Test cases.xlsx ← Excel test case file
│
└── README.md                           ← This file
```

> **Note:** The Excel file is resolved relative to the project root. You can place it anywhere and pass its path via `--excel`.

---

## ✅ Prerequisites

- **Python 3.10 or higher** — [Download](https://www.python.org/downloads/)
- **pip** (bundled with Python)
- An internet connection to reach `https://www.pixelssuite.com/chat-translator`

---

## 📦 Installation

### 1. Clone the repository

```bash
git clone https://github.com/udara780/chat-translator-test-automation.git
cd chat-translator-test-automation
```

### 2. (Recommended) Create a virtual environment

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python -m venv .venv
source .venv/bin/activate
```

### 3. Install Python dependencies

```bash
pip install playwright openpyxl
```

---

## 🌍 Installing Playwright Browsers

Playwright downloads and manages its own browser binaries. After installing the Python package, run:

```bash
playwright install chromium
```

> This downloads the Chromium binary used by the automation script. You only need to run this **once**.

To install all supported browsers (Chromium, Firefox, WebKit):

```bash
playwright install
```

---

## ▶️ Running the Tests

### Basic run (all defaults)

```bash
python test_automation/test_automation.py
```

The script will automatically locate `Assignment 1 - Test cases.xlsx` in the `test_automation/` directory and begin processing.

### Run with a custom Excel file

```bash
python test_automation/test_automation.py --excel "path/to/your/TestCases.xlsx"
```

### Run in headless mode (no visible browser window)

```bash
python test_automation/test_automation.py --headless
```

### Full example with all common options

```bash
python test_automation/test_automation.py \
  --excel "test_automation/Assignment 1 - Test cases.xlsx" \
  --sheet " Test cases" \
  --wait-ms 6000 \
  --retries 10 \
  --retry-wait-ms 1500 \
  --type-delay-ms 40 \
  --slow-mo-ms 200 \
  --save-every 5 \
  --headless
```

---

## ⚙️ Command-Line Parameters

| Parameter | Default | Description |
|---|---|---|
| `--excel` | Auto-detected | Path to the `.xlsx` test case file |
| `--sheet` | `" Test cases"` | Excel sheet (tab) name to read from |
| `--header-row` | `0` (auto-detect) | Row number of the header row (0 = auto-scan) |
| `--max-header-scan-rows` | `30` | How many rows to scan when auto-detecting the header |
| `--input-col` | Auto-detected | Name of the column containing Singlish input text |
| `--expected-col` | Auto-detected | Name of the column containing expected Sinhala output |
| `--actual-col` | `"Actual output"` | Name of the column where actual output will be written |
| `--status-col` | `"Status"` | Name of the column where PASS/FAIL status will be written |
| `--url` | PixelsSuite URL | Target website URL (overridable via `FRONTEND_URL` env var) |
| `--output` | Same as `--excel` | Path to save the results file (defaults to overwriting input) |
| `--save-every` | `0` (off) | Save the workbook after every N processed rows (0 = save only at end) |
| `--headless` | `false` | Run the browser in headless mode (no visible window) |
| `--wait-ms` | `5000` | Time (ms) to wait after clicking Transliterate before reading output |
| `--retries` | `8` | Number of times to retry reading the output if it appears empty |
| `--retry-wait-ms` | `1000` | Delay (ms) between each output-read retry |
| `--type-delay-ms` | `30` | Delay (ms) between each keystroke when typing input (simulates human typing) |
| `--timeout-ms` | `60000` | Global Playwright element wait timeout in milliseconds |
| `--slow-mo-ms` | `0` | Slows all browser actions by this many milliseconds (useful for debugging) |
| `--keep-open` | `false` | Keeps the browser open after all tests; press `CTRL+C` to exit |

---

## 📊 Excel Sheet Format

The script performs **fuzzy header detection**, so column names do not need to match exactly. Supported variants are listed below.

### Required Columns

| Column Role | Accepted Header Names |
|---|---|
| **Input** (Singlish text) | `Singlish`, `Input`, `Singlish Input`, `Test Input`, `Source`, `Sentence`, `Text` |
| **Expected Output** (Sinhala) | `Sinhala`, `Expected_Output`, `Expected Output`, `Expected output`, `Expected`, `Expected Sinhala` |

### Auto-Created Columns

If the following columns do not already exist, the script will **add them automatically**:

| Column Role | Default Name |
|---|---|
| **Actual Output** | `Actual output` |
| **Test Status** | `Status` |

### Sample Excel Layout

| TC ID | Singlish | Expected Output | Actual output | Status |
|---|---|---|---|---|
| TC-001 | mama | මම | මම | PASS |
| TC-002 | kohomada | කොහොමද | කොහොමද | PASS |
| TC-003 | ayubowan | ආයුබෝවන් | ආයුබෝවන් | PASS |
| TC-004 | giyada | ගියද | ගිය | FAIL |

> 💡 **Tip:** Rows with an empty input cell are automatically skipped. Rows without an expected output are marked `COLLECTED` instead of PASS/FAIL.

---

## 🔄 Sample Workflow

```
1. Prepare Excel file with Singlish inputs and expected Sinhala outputs
        │
        ▼
2. Run the automation script
   python test_automation/test_automation.py --save-every 5
        │
        ▼
3. Script launches Chromium and navigates to the Chat Translator
        │
        ▼
4. For each test row:
   a. Dismiss any cookie/consent overlays
   b. Type the Singlish input (with configurable delay)
   c. Click the "Transliterate" button
   d. Wait for the output textarea to update
   e. Retry if output is empty or unchanged
   f. Write actual output → Excel
   g. Compare with expected → Write PASS / FAIL / COLLECTED
        │
        ▼
5. Save the updated Excel file with all results
        │
        ▼
6. Review the annotated Excel report
```

---

## 📋 Expected Output

### Console output during a run

```
Starting Frontend-Only test with 20 rows...
Frontend loaded successfully.
Testing [Row 3]: mama
  -> PASS
Testing [Row 4]: kohomada
  -> PASS
Testing [Row 5]: giyada
  -> FAIL
...
Test completed. Results saved to test_automation\Assignment 1 - Test cases.xlsx
```

### Status values written to Excel

| Status | Meaning |
|---|---|
| `PASS` | Actual output exactly matches the expected output |
| `FAIL` | Actual output does not match the expected output |
| `COLLECTED` | No expected output was provided; actual output was captured for reference |
| `UI Error` | A Playwright interaction error occurred while processing this row |

---

## 🔧 Troubleshooting

### ❌ `playwright install` fails or browser not found

```bash
# Re-install Playwright and force browser download
pip install --upgrade playwright
playwright install chromium
```

### ❌ `File not found` error for the Excel file

Ensure the file path is correct and that the file is not currently open in Excel (Excel locks the file on Windows):

```bash
python test_automation/test_automation.py --excel "test_automation/Assignment 1 - Test cases.xlsx"
```

### ❌ Output column always shows empty / status always `FAIL`

Increase `--wait-ms` and `--retries` to give the translator more time to respond:

```bash
python test_automation/test_automation.py --wait-ms 8000 --retries 15
```

### ❌ Script says "Could not find Chat UI locators"

The website may have changed its HTML structure. Run in headed mode with slow-motion to inspect the page:

```bash
python test_automation/test_automation.py --slow-mo-ms 500 --keep-open
```

### ❌ Unicode/Sinhala characters appear garbled in the console

This is a Windows console encoding issue and does not affect the Excel output. The script automatically reconfigures `stdout` to UTF-8. If the issue persists, run:

```powershell
chcp 65001
python test_automation/test_automation.py
```

### ❌ "Permission denied" when saving the Excel file

Close the Excel file in Microsoft Excel before running the script, as Excel holds a file lock that prevents writing.

---

## 🚀 Future Improvements

- [ ] **Multi-language support** — extend to test other language pairs supported by the Chat Translator
- [ ] **Screenshot on failure** — automatically capture a screenshot when a test row fails for visual evidence
- [ ] **HTML test report** — generate a rich HTML report (e.g., via `pytest-html`) alongside the Excel output
- [ ] **CI/CD integration** — add a GitHub Actions workflow to run tests on every push
- [ ] **Parallel execution** — run multiple browser contexts in parallel to speed up large test suites
- [ ] **Configurable browser** — add support for Firefox and WebKit via a `--browser` CLI flag
- [ ] **CSV/Google Sheets input** — support reading test cases directly from Google Sheets or CSV files
- [ ] **Retry on network error** — detect network timeouts and automatically reload the page before retrying

---

## 🏁 Conclusion

This project demonstrates a **practical, production-style test automation framework** built around real-world constraints:

- Test data is fully separated from code, managed in Excel for easy collaboration
- The browser interaction layer is robust — handling overlays, merged cells, fuzzy headers, and flaky output timing
- Every behaviour is configurable from the command line, making the suite adaptable to different environments without touching source code

The framework is suitable as a foundation for larger test suites and can be extended with reporting, parallel execution, or CI/CD integration with minimal effort.

---

> **Author:** IT23265660 | **Assignment:** 1 | **Module:** Software Quality Assurance
>
> *Tested against [https://www.pixelssuite.com/chat-translator](https://www.pixelssuite.com/chat-translator)*

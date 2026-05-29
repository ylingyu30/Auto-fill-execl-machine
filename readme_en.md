**这是一个自动填表器**  
**This is an auto-fill Excel machine**

简体中文[readme.md](readme.md) file for details.

It can automatically select keywords from different word banks based on score values in specified cells (e.g., E3, E4) of an Excel sheet, fill them into corresponding cells (e.g., B5, B6, F3, F4), and support full-row data filling. It is suitable for batch scenarios such as generating personalized comments or student feedback.

## ✨ Features

- 📊 **Multi-subject support**: One Excel file can contain multiple worksheets (e.g., Chinese, Math, English), each with independently configurable filling rules.
- 🎯 **Conditional filling**: Automatically selects the appropriate word bank based on the numeric value in the score cell (e.g., E3). For example, use the “Excellent” word bank for scores ≥75, and the “Needs Improvement” word bank for scores <75.
- 🚫 **Invalid score handling**: If the score cell is empty or contains markers like “absent” or “none”, filling can be skipped (original content remains unchanged).
- 📥 **Batch processing**: Recursively scans all `.xlsx` files in the `data` folder and outputs them into the `output` folder while preserving the directory structure.
- 🔧 **Flexible configuration**: All rules are defined via a JSON configuration file – no code modification required.
- 📦 **Packaged as a standalone exe**: One-click packaging scripts (`install.bat` / `install.sh`) are provided to generate an `.exe` file that can run without a Python environment.

## 🚀 Quick Start

### 1. Prepare the environment (for running from source)
- Python 3.8 or higher (3.10+ recommended)
- Install dependency: `pip install openpyxl`

### 2. Prepare word bank files
Organize word banks by subject and cell. Each word bank is a UTF-8 encoded text file with one word or phrase per line.  
Example: `wordbanks/English/B5good.txt`

Full-row word banks are CSV files (comma separated), e.g., `wordbanks/English/A12-E12.csv`:

### 3. Write the configuration file `fill_config.json`
```json
{
    "input_dir": "data",
    "output_dir": "output",
    "subjects": [
        {
            "name": "English",
            "sheet_name": "English",
            "fill_rules": {
                "B5": {
                    "score_cell": "E3",
                    "thresholds": [75, 50],
                    "wordbanks": ["wordbanks/English/B5good.txt", "wordbanks/English/B5med.txt", "wordbanks/English/B5low.txt"],
                    "skip_on_invalid": true,
                    "invalid_markers": ["absent", "N/A", "none", "leave"]
                },
                "B6": {
                    "score_cell": "E4",
                    "thresholds": [80, 60],
                    "wordbanks": ["wordbanks/English/B6good.txt", "wordbanks/English/B6med.txt", "wordbanks/English/B6low.txt"],
                    "skip_on_invalid": true,
                    "invalid_markers": ["absent", "N/A", "none", "leave"]
                },
                "F3": {
                    "score_cell": "E3",
                    "thresholds": [75, 50],
                    "wordbanks": ["wordbanks/English/F3good.txt", "wordbanks/English/F3med.txt", "wordbanks/English/F3low.txt"],
                    "skip_on_invalid": true,
                    "invalid_markers": ["absent", "N/A", "none", "leave"]
                },
                "F4": {
                    "score_cell": "E4",
                    "thresholds": [80, 60],
                    "wordbanks": ["wordbanks/English/F4good.txt", "wordbanks/English/F4med.txt", "wordbanks/English/F4low.txt"],
                    "skip_on_invalid": true,
                    "invalid_markers": ["absent", "N/A", "none", "leave"]
                },
                "A12:E12": {
                    "file": "wordbanks/English/A12-E12.csv",
                    "delimiter": ","
                }
            }
        },
        {
            "name": "Math",
            "sheet_name": "Math",
            "fill_rules": {
                "B5": {
                    "score_cell": "E3",
                    "thresholds": [70, 40],
                    "wordbanks": ["wordbanks/Math/B5good.txt", "wordbanks/Math/B5med.txt", "wordbanks/Math/B5low.txt"]
                },
                "A12:E12": {
                    "file": "wordbanks/Math/A12-E12.csv",
                    "delimiter": ","
                }
            }
        }
    ]
}
```

### 4. Run the script
```bash
python fill_excel.py
```

## 📦 Package as EXE (no Python environment required)

The project provides packaging scripts `install.bat` (Windows) and `install.sh` (Linux/macOS).

### Windows
```cmd
install.bat fill_excel.py
```
Wait for completion; the generated `fill_excel.exe` is located in the `dist` folder.

### Linux/macOS
```bash
chmod +x install.sh
./install.sh fill_excel.py
```
The generated binary is in the `dist` directory.

After packaging, you can distribute the entire `output` directory together with the word bank folder – no need to install Python.

## 📂 Example directory structure

```
project root/
├─ fill_excel.py          # Main script
├─ fill_config.json       # Configuration file
├─ install.bat            # Windows packaging script
├─ install.sh             # Linux/macOS packaging script
├─ data/                  # Excel files to be processed
│  ├─ ZhangSan.xlsx
│  └─ LiSi.xlsx
├─ output/                # Output after processing (auto-created)
├─ wordbanks/             # Word bank folder
│  ├─ English/
│  │  ├─ B5good.txt
│  │  ├─ B5med.txt
│  │  ├─ B5low.txt
│  │  ├─ B6good.txt
│  │  ├─ ...
│  │  └─ A12-E12.csv
│  └─ Math/
│      └─ ...
```

## ⚙️ Configuration details

### General settings
- `input_dir`: Folder containing the original Excel files (scans all `.xlsx` recursively)
- `output_dir`: Output folder (preserves the same directory structure)
- `subjects`: List of subjects, each containing:
  - `name`: Subject name (used only for logging)
  - `sheet_name`: Worksheet name in the Excel file (must match exactly)
  - `fill_rules`: Dictionary of filling rules

### Filling rules
- **Single‑cell filling** (e.g., `"B5"`):
  - `score_cell`: Cell address from which to read the score (e.g., `"E3"`)
  - `thresholds`: Array of thresholds in descending order, e.g., `[75, 50]` means:
    - score ≥ 75 → use `wordbanks[0]`
    - 50 ≤ score < 75 → use `wordbanks[1]`
    - score < 50 → use `wordbanks[2]`
  - `wordbanks`: Array of word bank file paths corresponding to the threshold intervals (length = `len(thresholds)+1`)
  - `skip_on_invalid`: Whether to skip filling when the score is invalid (default `true`)
  - `invalid_markers`: List of keywords that identify an “invalid score” (e.g., `["absent", "N/A"]`). If the cell content contains any of these keywords, it is treated as invalid.
- **Full‑row filling** (e.g., `"A12:E12"`):
  - `file`: Path to a CSV file, each line represents the fields to be filled into the entire row
  - `delimiter`: Field delimiter, defaults to `,`

### Score validity logic
1. Cell is empty → invalid
2. Cell content contains any string from `invalid_markers` → invalid
3. Cell content cannot be converted to a number → invalid
4. Successfully converted to a number → valid, used for threshold comparison

## 💡 Frequently Asked Questions

### 1. The script reports `No module named 'openpyxl'`
- Dependency not installed: run `pip install openpyxl`
- If you have packaged it as an exe, make sure the `--hidden-import=openpyxl` parameter was used during packaging (the `install.bat` script already includes it by default).

### 2. Chinese characters appear garbled when opening a CSV file in Excel
- Save the word bank files as **UTF-8 with BOM** encoding (using Notepad or Notepad++).
- Alternatively, use Excel’s “Data → From Text/CSV” import and select UTF-8 encoding.

### 3. How to handle subject‑specific word banks?
- Use sub‑folders inside the `wordbanks` path, e.g., `wordbanks/Chinese/`, `wordbanks/Math/`.
- In the configuration file, specify the corresponding word bank paths for each subject’s `fill_rules`.

### 4. The packaged exe is flagged as virus or blocked
- Files packaged with PyInstaller are sometimes falsely detected by antivirus software – simply add an exception.
- You can use Nuitka as an alternative packager, which has a lower false‑positive rate.

## 📜 Open Source License

This project is open sourced under the **GPL 3.0** license. See the [LICENSE](LICENSE) file for details.

---

**Enjoy!** If you have any questions, feel free to open an issue.

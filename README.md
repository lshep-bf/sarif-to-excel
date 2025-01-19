# SARIF to Excel Formatter

**SARIF to Excel Formatter** is a Python-based tool designed to transform SARIF (Static Analysis Results Interchange Format) reports into clean, readable, and professionally formatted Excel files. This tool simplifies the analysis of static code analysis results, making it easier to interpret, share, and report findings.

---

## Features

- **Dynamic Column Resizing**:
  Automatically adjusts column widths based on the largest content for improved readability.

- **Text Wrapping for Large Fields**:
  Ensures fields like `Message` and `Details` are legible by wrapping text and setting optimal column widths.

- **Customizable Columns**:
  Allows you to include specific fields like `Severity`, `Path`, `Page`, `Line`, and more, with a user-friendly layout.

- **Tool-Agnostic Compatibility**:
  Works with any tool that outputs SARIF reports, including:
  - Qodana
  - CodeQL
  - SonarQube
  - ESLint
  - GitHub Code Scanning

- **Excel Table Formatting**:
  Outputs professional-grade Excel files with built-in table formatting for filtering and sorting.

---

## Requirements

Ensure the following dependencies are installed:

- Python 3.6+
- Required Python libraries:
  ```bash
  pip install pandas openpyxl
  ```

---

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/barkerbg001/sarif-to-excel.git
   ```

2. Navigate to the project directory:
   ```bash
   cd sarif-to-excel
   ```

3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

### Basic Command:
Run the script with your SARIF file:

```bash
python main.py
```

### Input:
- Provide the path to your SARIF file (e.g., `qodana.sarif.json`).

### Output:
- The formatted Excel file (`sarif_report.xlsx`) will be generated in the project directory.

---

## Example SARIF Report
Input (SARIF):
```json
{
  "version": "2.1.0",
  "runs": [
    {
      "results": [
        {
          "ruleId": "EXAMPLE_RULE",
          "message": { "text": "This is a test message." },
          "locations": [
            {
              "physicalLocation": {
                "artifactLocation": { "uri": "src/example.js" },
                "region": { "startLine": 42 }
              }
            }
          ],
          "level": "error"
        }
      ]
    }
  ]
}
```

Output (Excel):
| Severity | Message       | Details               | Path           | Page          | Line |
|----------|---------------|-----------------------|----------------|---------------|------|
| error    | EXAMPLE_RULE  | This is a test message.| src/example.js | example.js    | 42   |

---

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository.
2. Create a new branch:
   ```bash
   git checkout -b feature/your-feature
   ```
3. Commit your changes:
   ```bash
   git commit -m "Add your message here"
   ```
4. Push to your branch:
   ```bash
   git push origin feature/your-feature
   ```
5. Open a pull request.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Acknowledgments

- [SARIF Specification](https://sarifweb.azurewebsites.net/)
- Tools like [Qodana](https://www.jetbrains.com/qodana/), [CodeQL](https://codeql.github.com/), and [SonarQube](https://www.sonarqube.org/) for inspiring this project.

---

## Contact

For questions or suggestions, feel free to reach out:
- **GitHub**: [barkerbg001](https://github.com/barkerbg001)


# SARIF to Excel Formatter

A Python tool that converts SARIF (Static Analysis Results Interchange Format) reports into professionally formatted Excel spreadsheets with clickable hyperlinks, custom styling, and organized data.

**Note that this fork of the tool is meant to solve specific problems encountered with SARIF files that are create with the (excellent) `trivy`  command line tool. But it should still work with SARIF files created by other tools.**

---

## Features

- **Professional Excel Formatting**:
  - Clean table layout with thin black borders
  - Dark grey header row with white bold text
  - Filter dropdowns on all columns

- **Smart Data Processing**:
  - Normalizes severity levels (`note` → `Low`, `warning` → `Medium`, `error` → `High`)
  - Converts Aquasec URLs to NIST URLs automatically
  - Creates clickable hyperlinks from Markdown-formatted links

- **Dynamic Column Sizing**:
  - Auto-fits columns for `Path`, `Page`, and `Line`
  - Text wrapping for `Message` and `Details` columns

- **Sheet Naming**:
  - Names Excel sheets after the input SARIF filename for easy organization

- **Tool Compatibility**:
  Works with any SARIF-compliant tool (Trivy, Qodana, CodeQL, SonarQube, ESLint, GitHub Code Scanning)

---

## Requirements

- Python 3.6+
- Dependencies:
  ```bash
  pip install pandas openpyxl
  ```

---

## Installation

1. Clone this repository. Note that this is a fork of the original repo created by [barkerbg001](https://github.com/barkerbg001):
   ```bash
   git clone https://github.com/lshep-bf/sarif-to-excel.git
   cd sarif-to-excel
   ```

2. Install dependencies:
   ```bash
   pip install pandas openpyxl
   ```

---

## Usage

Run the tool with a SARIF file path as argument:

```bash
python main.py path/to/your-report.sarif
```

### Output:
- Excel file created in the same directory as the input file
- Filename matches the input (e.g., `your-report.sarif` → `your-report.xlsx`)
- Sheet tab named after the input file

### Example:
```bash
python main.py ecommerce-api.sarif
```
Creates `ecommerce-api.xlsx` with a sheet tab named `ecommerce-api`.

---

## Excel Output Structure

| Severity | Message | Details | Path | Page | Line |
|----------|---------|---------|------|------|------|
| Low/Medium/High | Rule ID | Description with clickable links | Full file path | Filename | Line number |

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
- **Owner of this fork**: [lshep-bf](https://github.com/lshep-bf)
- **Creator of the original repo**: [barkerbg001](https://github.com/barkerbg001)


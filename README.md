# ASVS Excel Generator

Small utility script that downloads the [OWASP Application Security Verification Standard](https://owasp.org/www-project-application-security-verification-standard/) (ASVS) CSV and generates a structured Excel workbook to assess and document ASVS requirements fulfilment. Each chapter is represented as a worksheet with fulfilment tracking per requirement, supported by per-chapter summary tables and an overall compliance summary chart.

The repository also provides pre-generated Excel workbooks for ASVS 4.0.3 and ASVS 5.0.0 for anyone who prefers to use the templates directly rather than generating them locally.

## Requirements
- Python 3.10+
- Dependencies from `requirements.txt`

Install:
```bash
pip install -r requirements.txt
```

## Usage

Default: generates ASVS 5.0.0 as `OWASP-ASVS-5.0.0.xlsx`.

```bash
python asvs.py
```

Version specification: generates ASVS 4.0.3 as `OWASP-ASVS-4.0.3.xlsx`.

```bash
python asvs.py --asvs-version 4
```

Output specification: writes the workbook to a provided path.

```bash
python asvs.py --output "out/OWASP-ASVS-5.0.0.xlsx"
```

Full list of options:
- `-a`, `--asvs-version`: ASVS version `4` or `5` (default: `5`)
- `-o`, `--output`: output `.xlsx` file path (default depends on version)
  - v4: `OWASP-ASVS-4.0.3.xlsx`
  - v5: `OWASP-ASVS-5.0.0.xlsx`


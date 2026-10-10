# Contributing

Thanks for your interest in improving IQC Barcode Scanner!

## Ground rules

- **Keep it a single file.** `iqc_scanner.py` is intentionally self-contained so it is easy to read, audit and package. Please do not split it into modules in a pull request without opening an issue first.
- **Small, focused pull requests** with a clear description are easier to review.
- **Never include real data**: no real barcodes, customer or project names, connection strings, credentials or screenshots with sensitive information. Use dummy data only.

## Reporting bugs

Open an issue and include:

1. What you did and what you expected
2. What happened instead (error text or screenshot)
3. Python version, OS, ODBC driver and SQL Server version
4. The relevant `config.xlsx` row (with sensitive values removed)

## Suggesting features

Open an issue describing the problem you want to solve before writing code.

## Development setup

```bash
git clone https://github.com/dineshnaidu10/iqc-barcode-scanner.git
cd iqc-barcode-scanner
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt
```

Then follow the Quick start in the README (database, `db_config.ini`, `config.xlsx`).

## Pull request checklist

- [ ] `python -m py_compile iqc_scanner.py` passes
- [ ] Tested with `config.sample.xlsx` and `samples/sample_barcodes.txt`
- [ ] README updated if behaviour or configuration changed
- [ ] No secrets or real data included

## Code style

PEP 8 where practical, matching the surrounding code. Keep UI colours in the `DARK` theme dictionary.

# AGENTS.md

## Cursor Cloud specific instructions

### Repository overview

This is a GitHub profile README repository (`Ansidrake/Ansidrake`). The `main` branch contains only `README.md`. Executable code lives on feature branches:

- **`cursor/icai-audit-templates-db13`** — Python script (`scripts/build_icai_audit_templates.py`) that generates ICAI audit documentation templates in `.docx` and `.xlsx` format. Dependencies are listed in `scripts/requirements-audit-templates.txt`.

### Running the audit template generator

```bash
pip install -r scripts/requirements-audit-templates.txt
python3 scripts/build_icai_audit_templates.py
```

Output goes to `templates/` in the repo root.

### Notes

- There are no lint configurations, test suites, build systems, or long-running services in this repository.
- The generated `templates/` directory is already committed on the feature branch; re-running the script regenerates the same files.
- Python 3.10+ is required (uses `pathlib`, type hints, and `match` not present but `f-strings` used extensively).

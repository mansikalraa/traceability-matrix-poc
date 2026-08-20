## Traceability Matrix

| Requirement | Hazard | Design Module | Test Case | Verification |
|---|---|---|---|---|
| [FR-001](https://github.com/mansikalraa/traceability-matrix-poc/issues/1) | [H-001](https://github.com/mansikalraa/traceability-matrix-poc/blob/main/docs/risk-management.md#h-001--invalid-dicom-upload) | File Upload | [TC-001](https://github.com/mansikalraa/traceability-matrix-poc/issues/5) | Pass |
| [FR-002](https://github.com/mansikalraa/traceability-matrix-poc/issues/2) | [H-002](https://github.com/mansikalraa/traceability-matrix-poc/blob/main/docs/risk-management.md#h-002--incorrect-ivd-measurement) | Measurement Engine | [TC-002](https://github.com/mansikalraa/traceability-matrix-poc/issues/6) | -- |
| [FR-003](https://github.com/mansikalraa/traceability-matrix-poc/issues/3) | [H-003](https://github.com/mansikalraa/traceability-matrix-poc/blob/main/docs/risk-management.md#h-003--incorrect-surgical-trajectory) | Trajectory Engine | [TC-003](https://github.com/mansikalraa/traceability-matrix-poc/issues/7) | -- |
| [FR-004](https://github.com/mansikalraa/traceability-matrix-poc/issues/4) | [H-004](https://github.com/mansikalraa/traceability-matrix-poc/blob/main/docs/risk-management.md#h-004--incorrect-navigation-export) | Navigation Export | -- | -- |

---

## Document Generation & Automation

The repository includes an automated script [`generate_matrix_report.py`](file:///Users/mansikalra/Desktop/traceability-matrix-poc/generate_matrix_report.py) to compile the matrix and documentation into Word (`.docx`) or PDF (`.pdf`) compliance reports.

### Modes of Execution

1. **Mode 1 (Summary & Audit Only - Steps 1 & 2)**:
   Generates a cover page with the Traceability Matrix table and an audit of incomplete rows or open issues.
   ```bash
   python3 generate_matrix_report.py --mode 1 --format docx --output summary.docx
   ```

2. **Mode 2 (Full Compliance Report - Steps 1 to 4)**:
   Executes all 4 steps: cover page, incomplete rows audit, stacked documentation (Use Cases -> Risk Management -> Test Cases) with latest PR review details (Author, Reviewers, Title, Description), and complete issue/PR hyperlinks.
   ```bash
   python3 generate_matrix_report.py --mode 2 --format pdf --output full_report.pdf
   ```

### Output Format Options
- `--format docx`: Generates a Microsoft Word document (default).
- `--format pdf`: Generates a PDF document.
- `--format both`: Generates both `.docx` and `.pdf` files simultaneously.

### Running via GitHub Actions

You can manually trigger document generation directly in GitHub:
1. Go to the **Actions** tab in your repository.
2. Select **Generate Traceability Matrix Report** workflow.
3. Click **Run workflow**, choose your desired **Execution Mode** (`1` or `2`) and **Output Format** (`both`, `docx`, `pdf`), and click **Run workflow**.
4. Download the generated `.docx` and `.pdf` files from the **Artifacts** section of the workflow run summary.



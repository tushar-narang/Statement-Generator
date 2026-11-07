# Statement Generator

Generates a monthly Excel statement from the latest workbook, keeps totals within ±10% of the previous recorded month, creates trend charts, and emails the Excel file and charts.

## Repository layout

- \`December2025.xlsx\`: original workbook template / fallback.
- \`history.csv\`: monthly totals and the filename of each generated workbook.
- \`generated_files/\`: generated workbooks and chart images.
- \`main.py\`: generator.
- \`.github/workflows/monthly-statement.yml\`: scheduled and manual GitHub Actions workflow.

## GitHub Actions setup

In **Settings → Secrets and variables → Actions**, add:

- \`EMAIL_ADDRESS\`: Gmail sender address.
- \`EMAIL_PASSWORD\`: Gmail App Password (not your normal Gmail password).
- \`RECEIVER_EMAIL\`: destination email address.

Enable 2-Step Verification on the Gmail account and create an App Password from https://myaccount.google.com/security.

The workflow runs at 18:00 India time on the last calendar day of each month. It also supports manual runs from **Actions → Generate Monthly Statement → Run workflow**. Manual runs are allowed on any day.

## Run locally

\`\`\`bash
python -m venv .venv
# macOS/Linux
source .venv/bin/activate
# Windows PowerShell
.venv\\Scripts\\Activate.ps1

pip install -r requirements.txt
python main.py
\`\`\`

Set \`OUTPUT_MONTH_YEAR\` to override the output month label, for example \`OUTPUT_MONTH_YEAR=November2026\`.

## Notes

- The most recent workbook is selected using the last \`file\` entry in \`history.csv\`, not file modification timestamps.
- The workbook's formulas are set to recalculate when opened in Excel. Totals are independently calculated and validated in Python before saving.
- The generator stops if required sheets or baseline totals are missing, or if any generated total falls outside the ±10% range.
- Email is skipped when the three email secrets are not configured.

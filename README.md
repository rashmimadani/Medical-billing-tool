# Cell-Optimal Lyfe Medical Solutions Billing Tool

This workflow helps a non-technical user paste a raw monthly staff log, validate the parsed calculations, append the monthly totals into an Income Tracker workbook, and generate client-ready invoices in both Excel and PDF formats.

## Features

- Paste raw monthly logs in plain text
- Parse patient names and service entries with flexible formatting
- Apply the 2026 billing rules, including the 45-minute grace rule
- Aggregate one patient into a single tracker row when multiple session types appear in the month
- Append rows to an existing Income Tracker workbook
- Generate individual invoices from a template workbook and export them to PDF
- Provide validation warnings for malformed rows instead of silently dropping data

## Run locally

1. Create a virtual environment if needed.
2. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

3. Launch the app:

   ```bash
   streamlit run app.py
   ```

4. Paste the monthly staff log and proceed through the workflow in the UI.

## Example log

```text
Ravi - 7 (30)(PT)
Abhinay - 1 (60) (PT)+ 7(30)(TPR)
Punit - 7 (30)(TPR)
Shilpa - 7 (30)(TPR)
Gautham - 2(60)(TPR)
Rashmi - 2(45)(PT)
Kush-2(45)(TPR)
George -4(30)(TPR)
Simran - 6 (30) (TPR)
Shanta - 2 (30)(PT)
Dr kamala -7(30)(TPR)
Bhanu- 5 (30)(TPR)
Col Sushil - 1(30)(PT)
Anu-1 (30)(PT)
Naina- 4(30)(TPR)
Shivaram-1(30)(TPR)
Dr.arun kumar- 2(30)(TPR)
```

## Notes

- The tool treats all 45-minute sessions as 30-minute billable sessions.
- A patient with multiple service types in the same month is combined into one tracker row and itemized separately on the invoice.
- If no tracker workbook or invoice template is available, the app creates demo files in the `data/` folder.

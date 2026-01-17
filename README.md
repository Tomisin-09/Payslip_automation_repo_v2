# Payslip Automation

This project automates the generation and delivery of employee payslips using **Excel as the source of truth**, **Python for orchestration**, and **Outlook for email delivery**.

It is designed to be:

- reliable
- auditable
- easy to understand
- easy to extend later (e.g. Microsoft Graph, ReportLab)

---

## What this does (high level)

For each pay period, the system:

1. Reads employee data from an Excel **Data Source** sheet
2. Uses an Excel **Payslip template** driven by a selector cell
3. Generates one `.xlsx` file per employee
4. Uses **Excel itself** to export a pixel-perfect PDF
5. Emails the PDF to each employee via **Outlook**
6. Can be scheduled automatically (Windows Task Scheduler)

No payroll logic is reimplemented in Python — **Excel handles all calculations**.

---

## Key design idea (important)

The Payslip sheet is controlled by **one named cell** (`EmployeeRef`).

When that value changes:

- Excel formulas pull the correct employee data
- totals recalculate automatically

Python simply:

- sets that value
- saves a copy
- asks Excel to export it

This keeps the system simple, auditable, and robust.

---

## Repository structure

```text
payslip-automation/
├─ config/
│  └─ settings.yml              # All configuration (paths, dates, email text)
│
├─ data/
│  └─ input/
│     └─ Test PayslipPDF.xlsx   # Master workbook (Data Source + Payslip)
│
├─ output/
│  └─ <period>/                # Generated per run (e.g. 2026-01)
│     ├─ xlsx/
│     ├─ pdf/
│     └─ logs/
│
├─ src/
│  ├─ main.py                  # Orchestrates the entire process
│  │
│  ├─ preflight.py             # Dependency & OS checks (fails early, clearly)
│  │
│  ├─ utils/
│  │  └─ date_resolver.py
│  │
│  ├─ io/
│  │  ├─ load_data.py
│  │  └─ template_writer.py
│  │
│  ├─ pdf/
│  │  └─ excel_pdf_exporter.py # Excel → PDF (Windows only)
│  │
│  └─ email/
│     └─ outlook_sender.py     # Outlook email sender (Windows only)
│
├─ requirements.txt
├─ run.sh                      # macOS/Linux one-command runner
├─ run.ps1                     # Windows PowerShell one-command runner
├─ run.bat                     # Double-click runner (avoids the user having to open the command line/terminal)
└─ README.md                   # Contains all the documentation relating to this process
```

---

## Requirements

### macOS (development only)

Used for:

- XLSX generation
- logic validation
- dry runs

**PDF export and email sending are intentionally skipped.**

**Required:**

- Python 3.9+

---

### Windows (production)

Used for:

- full XLSX → PDF → Outlook email flow

**Required:**

- Python 3.9+
- Microsoft Excel (desktop)
- Microsoft Outlook (configured account)

---

## Setup & running (recommended approach)

> **Important principle**  
> Always run pip via the Python interpreter (`python -m pip`).  
> This avoids PATH issues across macOS and Windows.

---

## macOS: first-time setup & run

```bash
cd "/path/to/payslip-automation"

python3 -m venv .venv
source .venv/bin/activate

python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

python3 src/main.py
```

### macOS behaviour

- XLSX files are generated
- PDF export is skipped
- Emails are not sent
- No errors are raised for Windows-only features

---

## Windows: first-time setup & run (PowerShell)

```powershell
cd "C:\path\to\payslip-automation"

py -m venv .venv
.\.venv\Scripts\Activate.ps1

py -m pip install --upgrade pip
py -m pip install -r requirements.txt

py src\main.py
```

### Windows behaviour

- XLSX files are generated
- PDFs are exported via Excel
- Emails are sent via Outlook

---

## One-command run scripts (recommended)

Once the repo is cloned, you can simply run:

### macOS / Linux

```bash
./run.sh
```

### Windows (PowerShell)

```powershell
.\run.ps1
```

### Windows (.bat file)

```
Simply double click the run.bat file and the process should run
```

These scripts:

- checks that python is installed on the system
- create the virtual environment if missing
- install dependencies
- run the program safely

---

## Configuration

All configuration lives in:

```text
config/settings.yml
```

You can control:

- pay period calculation (automatic or manual)
- input workbook location
- output folder structure
- PDF export behaviour
- email subject and body text

No code changes are required for normal operation.

---

## Testing safely (strongly recommended)

Before sending real emails:

1. In `settings.yml`:

   ```yaml
   email:
     enabled: true
   ```

2. In `outlook_sender.py`, temporarily set:

   ```python
   OutlookEmailSender(send_mode="display")
   ```

3. Run the script and review draft emails

4. Switch back to:

   ```python
   OutlookEmailSender(send_mode="send")
   ```

---

## Scheduling (Windows only)

Designed to be run via **Windows Task Scheduler**:

- Monthly trigger (e.g. 25th of each month)
- Action: run `py src\main.py`
- Start in: project root
- Run whether user is logged in or not

---

## Why Excel is used for PDF generation

Using Excel itself to export PDFs ensures:

- exact formatting
- consistent layout
- no font issues
- no reimplementation of payroll logic

This is intentional and avoids common automation pitfalls.

---

## Future enhancements

This architecture cleanly supports:

- Microsoft Graph email (headless)
- ReportLab PDF generation
- retry & failure isolation
- admin summary reports
- HTML email templates

---

## Summary

This project prioritises:

- correctness over cleverness
- clarity over abstraction
- Excel as the authority for payroll calculations

It is a strong, maintainable MVP designed to scale safely.

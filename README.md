# EXR Owner Financials Extractor — v3.0

Two ways to run the same extraction: command-line script or browser-based webapp.

---

## Project Structure

```
owner-financial-extractor/
│
│   ── Core Logic (shared) ──
├── extractor_core.py           <-- ALL extraction + output logic lives here
│
│   ── Command-Line Workflow (unchanged) ──
├── extract_owner_financials.py <-- CLI script (thin wrapper around core)
├── run.bat                     <-- double-click to run CLI
├── input/                      <-- drop .xlsx files here for CLI
├── output/                     <-- CLI writes output here
├── completed/                  <-- archive workflow
├── archive/                    <-- archive workflow
├── move_processed_files.py     <-- archive script
├── archive_files.bat           <-- archive batch file
│
│   ── Web App Workflow (new) ──
├── app.py                      <-- Streamlit webapp
├── run_webapp.bat              <-- double-click to launch webapp
│
│   ── Config ──
├── requirements.txt            <-- openpyxl + streamlit
```

---

## Setup (One Time)

Open a terminal in the project folder and run:

```
pip install -r requirements.txt
```

This installs both `openpyxl` (you already have this) and `streamlit` (new).

---

## Workflow 1: Command-Line Script (Unchanged)

This works exactly as before.

1. Drop `.xlsx` files into `input/`
2. Double-click `run.bat` (or run `python extract_owner_financials.py`)
3. Enter a property name when prompted
4. Output appears in `output/`

Nothing about this workflow has changed.

---

## Workflow 2: Streamlit Web App (New)

1. Double-click `run_webapp.bat` (or run `streamlit run app.py`)
2. Your browser opens to `http://localhost:8501`
3. Upload an `.xlsx` file
4. Enter a property name
5. Click **Extract Data**
6. Click **Download Datapack** to save the output

The webapp produces the exact same output workbook as the CLI script.

To stop the webapp, press `Ctrl+C` in the terminal window.

---

## How the Code is Organized

| File | What It Does | Who Calls It |
|------|-------------|--------------|
| `extractor_core.py` | All extraction logic, all output writers, the `process_workbook()` function | Both CLI and webapp |
| `extract_owner_financials.py` | Scans `input/` folder, prompts in console, saves to `output/` | You, via `run.bat` |
| `app.py` | File upload in browser, calls `process_workbook()`, offers download | You, via `run_webapp.bat` |

The key design principle: **`extractor_core.py` never prints to console and never reads from keyboard.** It takes inputs as arguments and returns results as data. This makes it usable from any frontend — CLI, webapp, or a future hosted API.

---

## Future: Hosted Internal App

When you're ready to host this for your team, the path is:

1. Deploy `extractor_core.py` + `app.py` to a cloud service (Streamlit Cloud, an EC2 instance, or a Docker container)
2. No changes to the core logic
3. Add authentication if needed (Streamlit has built-in auth for Streamlit Cloud)
4. Your CLI workflow at home continues to work unchanged

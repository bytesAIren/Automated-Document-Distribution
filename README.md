# Automated Document Distribution
### Excel + Google Apps Script Integration
#### Learning one mess at a time™

---

## What this is

A small but fully working automation that takes boring repetitive work — copying names, generating PDFs, and sending dozens of emails — and makes Excel and Gmail do it *for you*.

Built by a professional from a commercial back-office background with **zero formal programming training**, but a lot of patience, curiosity, and an LLM on speed dial.

---

## What it does

1. **Reads** customer data from an Excel / Google Sheet
2. **Generates** personalised PDFs by injecting data into Word templates (V1, V2, V3)
3. **Sends** each PDF as an individual email via Gmail
4. **Logs** every action back into the Sheet — so nothing gets sent twice

---

## How it works

```
Excel Sheet  →  VBA macro opens Word template
             →  replaces placeholders with row data
             →  exports PDF to local folder
             →  writes canonical filename to column G  ← source of truth
             →  marks row as Done in column F

Google Drive  ←  you upload the generated PDFs manually

GAS script   →  reads PDF filename from column G
             →  finds the PDF in Google Drive by exact name
             →  sends email with attachment
             →  marks row as Sent in column H
```

Column G is the **shared contract** between the two scripts — the exact filename written by VBA is the one GAS will search for in Drive. No guesswork, no mismatch.

---

## Tech stack

| Tool | Role |
|---|---|
| Microsoft Excel + VBA | Data parsing and local PDF generation |
| Microsoft Word | Document layout with `<<PLACEHOLDER>>` tags |
| Google Apps Script | Gmail API integration and email dispatch |
| Google Drive | Cloud storage for generated PDFs |
| Google Sheets | Shared data source and status log |

---

## Repository structure

```
/
├── demo/
│   ├── demo_data.xlsx          # Sample Excel workbook with example rows
│   └── V1.docx                 # Sample Word template with placeholders
├── output_example/
│   ├── sheet_screenshot.png    # How the Sheet looks after running both scripts
│   └── generated_pdfs/         # Example output PDFs
├── src/
│   ├── excel_vba_snippet.bas   # VBA macro — generates PDFs from Word templates
│   └── script_gmail.gs         # Apps Script — sends emails with Drive attachments
└── README.md
```

---

## Sheet structure

Both scripts share the same Google Sheet. **Column order matters.**

| Col | Field | Written by |
|---|---|---|
| A | Company Name | You |
| B | Record Code | You |
| C | Version (`V1` / `V2` / `V3`) | You |
| D | Email address | You |
| E | Date | You |
| F | VBA Status | Excel macro |
| G | **PDF Filename** ← source of truth | Excel macro |
| H | Email Status | Apps Script |

> **Do not edit columns F, G, or H manually.** They are written automatically by the scripts. See `demo/demo_data.xlsx` for a working example.

---

## Word template setup

Each template (`V1.docx`, `V2.docx`, `V3.docx`) must contain these exact placeholder tags as plain text:

| Placeholder | Replaced with |
|---|---|
| `<<CODE>>` | Record Code (col B) |
| `<<COMPANY>>` | Company Name (col A) |
| `<<EMAIL>>` | Email address (col D) |
| `<<DATE>>` | Date (col E) |

Placeholders are replaced everywhere in the document: **body, headers, footers, and text boxes**. See `demo/V1.docx` for a ready-to-use example.

---

## Prerequisites

**On the Excel / Windows side:**
- Microsoft Excel (any recent version)
- Microsoft Word (must be installed — VBA automates it in the background)
- Template files (`V1.docx`, `V2.docx`, `V3.docx`) in the **same folder** as the Excel workbook

**On the Google side:**
- A Google account with access to Gmail, Google Sheets, and Google Drive
- Access to [Google Apps Script](https://script.google.com)

---

## Setup — step by step

### Part 1 — Prepare your data and templates

1. Open `demo/demo_data.xlsx` as a starting point, or create your own spreadsheet with the column layout shown above.
2. Fill in your data from **row 2** onwards (row 1 = headers). Columns A–E only — F, G, H will be filled by the scripts.
3. Use `demo/V1.docx` as a template reference, or create your own Word files with the `<<PLACEHOLDER>>` tags listed above.
4. Place all template files in the **same folder** as your Excel workbook.

---

### Part 2 — Run the VBA macro (Excel → PDF)

1. Open your Excel workbook and press `Alt + F11` to open the VBA editor.
2. Go to **Insert → Module** and paste the contents of `src/excel_vba_snippet.bas`.
3. Close the editor, then press `Alt + F8`, select `GeneratePDFs`, and click **Run**.
4. The macro will:
   - Open each Word template, replace placeholders, and export a PDF per row
   - Save PDFs into a `Generated_PDFs\` subfolder (created automatically)
   - Write the exact PDF filename into **column G**
   - Log `Done - yyyy-mm-dd hh:mm` (or an error message) into **column F**

---

### Part 3 — Upload PDFs to Google Drive

1. Go to [drive.google.com](https://drive.google.com) and create a dedicated folder (e.g. `PDF_Dispatch`).
2. Upload all files from your local `Generated_PDFs\` folder into it.
3. Copy the folder ID from its URL:
   ```
   https://drive.google.com/drive/folders/THIS_IS_YOUR_FOLDER_ID
   ```

---

### Part 4 — Run the GAS script (Drive → Gmail)

1. Open your Google Sheet and go to **Extensions → Apps Script**.
2. Delete any existing code and paste the contents of `src/script_gmail.gs`.
3. At the top of the script, fill in your IDs:
   ```javascript
   const SHEET_ID  = 'YOUR_SHEET_ID_HERE';   // from the Sheet URL
   const FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';
   ```
   > Your Sheet ID is in its URL: `https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_SHEET_ID/edit`
4. Click **Save**, then **Run → sendEmailsWithAttachments**.
5. On the first run, Google will ask you to authorise access to Gmail and Drive — follow the prompts.
6. The script will:
   - Read the PDF filename from **column G** and search for it in Drive
   - Send the email with the PDF attached
   - Log `Sent - yyyy-mm-dd hh:mm` (or an error) into **column H**

Check **View → Logs** in the Apps Script editor to monitor progress in real time.

---

## Re-running safely

Both scripts are idempotent — safe to re-run multiple times without duplicating work:

- **VBA**: skips any row where column F already contains `Done`
- **GAS**: skips any row where column H already starts with `Sent`

To reprocess a specific row, clear its status cell (F or H) and run again.

---

## Quota and limits

| Limit | Value |
|---|---|
| Gmail free account | ~100 emails / day |
| Gmail Workspace account | ~1,500 emails / day |
| Script safety cap (`MAX_EMAILS_PER_RUN`) | 90 (configurable in `script_gmail.gs`) |

If you hit the daily cap mid-run, the script stops and logs a warning. Re-run the next day — already-sent rows are skipped automatically.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| PDF not generated | Word template not found | Check template files are in the same folder as the workbook |
| Placeholder not replaced | Typo in the Word file | Tags must be exactly `<<CODE>>`, `<<COMPANY>>`, `<<EMAIL>>`, `<<DATE>>` |
| Column G is empty after running VBA | Macro stopped on an earlier error | Check column F for error messages on each row |
| GAS reports "PDF not found in Drive" | File not uploaded, or wrong folder ID | Verify the file exists in the Drive folder and `FOLDER_ID` is correct |
| "Authorisation required" error in GAS | First run only | Click through the Google OAuth permission prompts and run again |
| Email sent but no attachment | PDF was not in Drive at send time | Upload the file, clear column H for that row, and re-run |

---

## What I learned

- **AI is a co-pilot**: not magic, but a very patient mentor if you know how to ask the right questions
- **Debugging is detective work**: 80% logic, 20% persistence
- **Empowerment**: you don't need a CS degree to build tools that save you hours — you just need a problem worth solving

---

## License

Do whatever you want with it. Just don't send spam, keep the curiosity alive, and please credit *Learning one mess at a time™*.

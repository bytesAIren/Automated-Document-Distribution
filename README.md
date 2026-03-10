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

## Tech stack

| Tool | Role |
|---|---|
| Microsoft Excel + VBA | Data parsing and local PDF generation |
| Word (.docx templates) | Document layout with `<<PLACEHOLDER>>` tags |
| Google Apps Script | Gmail API integration and email dispatch |
| Google Drive | Cloud storage for generated PDFs |
| Google Sheets | Shared data source and status log |

---

## Repository structure

```
/
├── src/
│   ├── excel_vba_fixed.bas       # VBA macro — generates PDFs from Word templates
│   └── script_gmail_fixed.gs     # Apps Script — sends emails with Drive attachments
├── templates/                    # Put your V1.docx / V2.docx / V3.docx here
├── output_example/               # Example of a generated PDF
└── README.md
```

---

## Sheet structure

Both scripts share the same Google Sheet. **Column order matters.**

| Col | Field | Written by |
|---|---|---|
| A | Company Name | You |
| B | Record Code | You |
| C | Version (V1 / V2 / V3) | You |
| D | Email address | You |
| E | Date | You |
| F | VBA Status | Excel macro |
| G | **PDF Filename** ← source of truth | Excel macro |
| H | Email Status | Apps Script |

> **Why column G matters:** the VBA macro sanitises the company name (removes illegal characters) before building the PDF filename, then writes the exact filename to column G. The Apps Script reads column G directly — no guesswork, no mismatch.

---

## Prerequisites

### On the Excel / Windows side
- Microsoft Excel (any recent version)
- Microsoft Word (must be installed — VBA automates it in the background)
- The three template files (`V1.docx`, `V2.docx`, `V3.docx`) in the **same folder** as the Excel workbook

### On the Google side
- A Google account
- A Google Sheet (can be a copy of your Excel file, uploaded to Drive)
- A Google Drive folder where the PDFs will be uploaded
- Access to [Google Apps Script](https://script.google.com)

---

## Setup — step by step

### Part 1 — Prepare your Word templates

1. Create up to three Word documents: `V1.docx`, `V2.docx`, `V3.docx`
2. Inside each document, place these placeholder tags wherever you want data to appear:

| Placeholder | Replaced with |
|---|---|
| `<<CODE>>` | Record Code (col B) |
| `<<COMPANY>>` | Company Name (col A) |
| `<<EMAIL>>` | Email address (col D) |
| `<<DATE>>` | Date (col E) |

3. Placeholders work in the **body, headers, footers, and text boxes**
4. Save the templates in the same folder as your Excel workbook

---

### Part 2 — Run the VBA macro (Excel → PDF)

1. Open your Excel workbook and make sure your data starts on **row 2** (row 1 = headers)
2. Press `Alt + F11` to open the VBA editor
3. Go to **Insert → Module** and paste the contents of `excel_vba_fixed.bas`
4. Close the editor and press `Alt + F8`, select `GeneratePDFs`, click **Run**
5. The macro will:
   - Generate one PDF per row into a `Generated_PDFs\` subfolder
   - Write the exact PDF filename into **column G**
   - Log the result (`Done` or `ERROR - ...`) into **column F**

---

### Part 3 — Upload PDFs to Google Drive

1. Go to [drive.google.com](https://drive.google.com) and create a dedicated folder (e.g. `PDF_Distribution`)
2. Upload all files from your local `Generated_PDFs\` folder into it
3. Open the folder and look at its URL:
   ```
   https://drive.google.com/drive/folders/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
   ```
   Copy the long ID at the end — this is your **FOLDER_ID**

---

### Part 4 — Set up Google Apps Script (Google → Email)

1. Open your Google Sheet
2. Get your **SHEET_ID** from its URL:
   ```
   https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit
   ```
3. Go to **Extensions → Apps Script**
4. Delete any existing code and paste the contents of `script_gmail_fixed.gs`
5. At the top of the script, fill in your IDs:
   ```javascript
   const SHEET_ID  = 'YOUR_SHEET_ID_HERE';
   const FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';
   ```
6. Click **Save**, then click **Run** on `sendEmailsWithAttachments`
7. On the first run, Google will ask you to authorise access to Gmail and Drive — accept
8. The script will:
   - Read column G for the exact PDF filename to look up in Drive
   - Send the email with the PDF attached
   - Log `Sent - yyyy-mm-dd hh:mm` (or an error) into **column H**

---

## How the two scripts stay in sync

```
Excel VBA                          Google Apps Script
─────────────────────────────      ──────────────────────────────
Reads:   cols A, B, C, D, E       Reads:  cols B, D, G
Writes:  col F (VBA status)       Writes: col H (email status)
         col G (PDF filename) ──────────► used as Drive search key
```

Column G is the **single source of truth** for the PDF filename. The VBA macro handles character sanitisation (e.g. `O'Brien & Co.` → `O-Brien---Co`); the Apps Script never has to replicate that logic.

---

## Re-running safely

Both scripts are **idempotent** — you can run them multiple times without duplicating work:

- **VBA**: skips any row where column F already contains `Done`
- **Apps Script**: skips any row where column H already starts with `Sent`

To reprocess a specific row, simply clear its status cell (F or H) and run again.

---

## Quota and limits

| Limit | Value |
|---|---|
| Gmail free account | ~100 emails/day |
| Gmail Workspace account | ~1,500 emails/day |
| Script safety cap (`MAX_EMAILS_PER_RUN`) | 90 (configurable) |

If you hit the cap mid-run, the script stops and logs a warning. Re-run the next day — already-sent rows are skipped automatically.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| PDF not found in Drive | Filename mismatch | Check col G matches the actual file in Drive |
| Word opens visibly during macro | `wdApp.Visible` set to True | Set it back to `False` in the VBA |
| "Authorization required" in GAS | First run, permissions not granted yet | Click through the OAuth flow |
| Email sent but no attachment | PDF wasn't uploaded to Drive | Upload the file and re-run (clear col H first) |
| Template placeholders not replaced | Placeholder typo in Word | Make sure tags are exactly `<<CODE>>` etc. |

---

## What I learned

- **AI is a co-pilot**: not magic, but a very patient mentor if you know how to ask the right questions
- **Debugging is detective work**: 80% logic, 20% persistence
- **Empowerment**: you don't need a CS degree to build tools that save you hours — you just need a problem worth solving

---

## License

Do whatever you want with it. Just don't send spam, keep the curiosity alive, and please credit *Learning one mess at a time™*.

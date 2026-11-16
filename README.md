# PDF Studio

An offline, all-in-one PDF toolkit for Windows, built with Python, Tkinter and PyMuPDF. It has a modern dashboard UI, user sign-in, a 90-day free trial, and a dashboard that shows how many files you have converted by file type. All processing runs locally; your files never leave your PC.

## Features

| Category | Tools |
|---|---|
| Organize | Merge PDF, Split PDF (ranges / every N pages / every page), Remove pages, Extract pages |
| Optimize | Compress PDF (extreme / recommended / low), Repair PDF |
| Convert to PDF | JPG to PDF, Word to PDF, PowerPoint to PDF, Excel to PDF, HTML to PDF (file or URL) |
| Convert from PDF | PDF to JPG/PNG, PDF to Word, PDF to Excel |
| Edit | Rotate, Add page numbers, Add watermark (text or image), Crop |
| Security | Unlock PDF, Protect PDF (AES-256, print / copy / edit permissions) |

Also included:

- **Authentication:** create an account, sign in and sign out. Passwords are stored as salted PBKDF2 hashes.
- **90-day trial:** starts when the first account is created and applies to the whole PC. The sidebar and dashboard show the days left. When the trial ends, the tools are blocked and the dashboard stays viewable.
- **Dashboard:** total files converted, most produced file type, number of file types, trial days left, a bar chart of files by type, and recent activity. Statistics are kept per user.

## Requirements

- Windows 10 or 11
- Python 3.10 or newer (with Tkinter, included in the standard python.org installer)
- For some tools:
  - **Word / PowerPoint / Excel to PDF:** Microsoft Office, or LibreOffice as a fallback
  - **HTML to PDF:** Microsoft Edge or Google Chrome

## Install and run

```bash
pip install -r requirements.txt
python main.py
```

On first launch, create an account. Output files are saved to `Documents\PDF Studio Output` unless you change the "Save to" folder on a tool page.

## Build a standalone .exe

```bash
pyinstaller --noconfirm --windowed --name PDFStudio --collect-all pdf2docx --hidden-import win32com.client main.py
```

The result is in `dist\PDFStudio\`. Distribute the whole folder.

## Project files

```
PDFStudio/
├── main.py            # the whole application (UI, tools, accounts, trial, stats)
├── requirements.txt   # Python dependencies
└── README.md
```

## Where data is stored

Everything is kept in `%APPDATA%\PDFStudio`:

| File | Contents |
|---|---|
| `users.json` | Accounts (name, email, salted password hash) |
| `license.json` | Signed trial start date (a second copy is kept in the Windows registry under `HKCU\Software\PDFStudio`) |
| `stats.json` | Per-user conversion counts and recent activity |

To reset a test machine, delete this folder and the registry key above.

## Before you ship

Change the `KEY` value inside the accounts section of `main.py` (search for `change-this-secret-before-you-ship`) to your own secret. It signs the trial date so it cannot be edited by hand.

Sign-in and the trial are enforced locally on each PC, so they suit a trial or demo build but are not a server-backed licence. For paid activation keys or an online licence check, you would need a small backend.

## Usage notes

- **Page ranges:** use forms like `1-3,5,8-`, or `all`.
- **Split ranges:** each comma-separated range becomes its own file, for example `1-3,4-6`.
- **Compress** re-encodes large images, so results depend on how image-heavy the PDF is. If no saving is possible, the original is copied unchanged.
- **PDF to Word** works best on text-based PDFs. Scanned PDFs need OCR first (not yet included).
- **PDF to Excel** extracts detected tables into one sheet per table. If none are found, the page text is written to a single sheet.
- Password-protected PDFs must be unlocked first with the Unlock PDF tool.

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` |
| Office to PDF fails | Install Microsoft Office or LibreOffice, and close any open Word/Excel/PowerPoint windows |
| HTML to PDF fails | Install Microsoft Edge or Google Chrome |
| "Trial expired" message | The 90 days have passed; a licence is needed to continue |

## Not yet included

OCR, Sign PDF, Redact, Edit PDF, Compare PDF, PDF to PowerPoint, PDF to PDF/A, and the thumbnail-based Organize screen.

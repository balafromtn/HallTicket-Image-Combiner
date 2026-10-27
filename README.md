# GACS HallTicket Image Combiner

A local Streamlit app for generating hall-ticket HTML and PDF files with student photographs. It supports single-student generation and bulk generation from an Excel/CSV roster.

> **Important:** This tool creates local copies and does not modify the college portal. Use it only for records you are authorized to access. Student register numbers, dates of birth, photographs, and generated files are personal information; keep them private and verify every PDF before use.

## Features

- **Single Hall Ticket:** enter a register number and date of birth, upload a photo, and generate HTML/PDF files.
- **Multiple Hall Tickets:** upload an `.xlsx` or `.csv` roster and multiple student photographs.
- Match photos by register number; filenames can be like `24UCS250904.jpg` or `24-ucs-250904.jpeg`.
- Show progress and per-student status during bulk processing.
- Download the individual PDF/HTML, a bulk CSV report, or a ZIP of generated files.
- Keep output files in the local `generated/` directory.

## Requirements

- Python 3.10 or newer
- Google Chrome installed
- Internet access to reach the portal and download a compatible ChromeDriver on first run

## Setup on Windows

Open PowerShell or Command Prompt in the project folder.

1. Create a virtual environment:

   ```powershell
   py -m venv .venv
   ```

2. Activate it in PowerShell:

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

   Or use Command Prompt:

   ```bat
   .venv\Scripts\activate.bat
   ```

3. Install dependencies:

   ```powershell
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

4. Start the app:

   ```powershell
   streamlit run app.py
   ```

Streamlit should open your browser. If it does not, open the local address printed in the terminal, usually `http://localhost:8501`.

## Single hall-ticket workflow

1. Choose **Single Hall Ticket**.
2. Enter the register number and select the date of birth.
3. Upload the student's photo.
4. Click **Generate hall ticket**.
5. Download the PDF and HTML, then inspect the PDF's layout and photo position.

## Bulk workflow

1. Choose **Multiple Hall Tickets**.
2. Upload the Excel (`.xlsx`) or CSV roster.
3. Upload the corresponding student photo files together.
4. Click **Start bulk generation** and keep the app open until processing finishes.
5. Review the status table, then download the CSV report and ZIP.

### Spreadsheet format

The first worksheet or CSV header row must contain a register-number column and a date-of-birth column. Supported header variants include `regno`, `register number`, `registration number`, `dob`, `date of birth`, and `birthdate`; capitalization and punctuation are ignored.

Example:

| regno | dob |
|---|---|
| 24UCS250904 | 2005-04-12 |
| 24UCS250905 | 2005-07-23 |

Recognized DOB formats include `YYYY-MM-DD`, `DD-MM-YYYY`, `DD/MM/YYYY`, `MM/DD/YYYY`, and common month-name formats. Excel date cells are also supported.

**Tip:** Keep register numbers stored as text in Excel if they contain leading zeroes.

### Photo filenames

The filename without its extension should match the student's register number. Matching ignores punctuation and capitalization. Supported formats: PNG, JPG/JPEG, WEBP, and BMP.

## Output structure

```text
generated/
├── html/
│   └── hallticket_<register_number>.html
├── pdf/
│   └── hallticket_<register_number>.pdf
├── bulk_report.csv
└── halltickets_bundle.zip
```

Generated files are excluded from Git because they may contain student personal information.

## Troubleshooting

- **Chrome/ChromeDriver error:** Install or update Google Chrome. `webdriver-manager` downloads a compatible driver when needed; the first run requires internet access.
- **Portal login or timeout:** Check the internet connection, date of birth, and portal availability.
- **Photo not found in bulk:** Make sure the photo filename matches the register number.
- **Photo placement looks wrong:** Open and verify the generated PDF before distributing it. Portal layout changes may affect rendering.
- **Large bulk jobs:** The app opens a fresh browser session for each student. This is more isolated but can take time.

## Project structure

```text
GACS-HallTicket-Image-Combiner/
├── app.py
├── single_hallticket.py
├── bulk_hallticket.py
├── requirements.txt
├── README.md
├── packages.txt
├── .gitignore
└── .streamlit/
    └── config.toml
```

## License

This Project Comes under MIT License.
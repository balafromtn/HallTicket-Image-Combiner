# bulk_hallticket.py
"""
Bulk Hall-Ticket HTML + PDF Generator

Input:
    Excel (.xlsx) or CSV with regno and dob columns.
    A folder containing student photos named by register number.

Output:
    generated/html/hallticket_<register_number>.html
    generated/pdf/hallticket_<register_number>.pdf
    generated/bulk_report.csv

Run:
    python hallticket_scraper.py
"""

from __future__ import annotations

import base64
import csv
import mimetypes
import time
import re
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, InvalidSessionIdException
from webdriver_manager.chrome import ChromeDriverManager


PORTAL_URL = "https://slmgaccoe.edecampus.com/student/index.php"
HALLTICKET_URL = "https://slmgaccoe.edecampus.com/student/Hallticket"
PRINT_URL = "https://slmgaccoe.edecampus.com/student/HallticketPrint.php"
ASSET_BASE_URL = "https://slmgaccoe.edecampus.com/student/"
WAIT_SECONDS = 25

HERE = Path(__file__).resolve().parent

IMAGE_EXTENSIONS = {
    ".png", ".jpg", ".jpeg", ".webp", ".bmp"
}


# ---------------------------------------------------------
# File and folder selection
# ---------------------------------------------------------

def choose_input_file() -> Path:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        filename = filedialog.askopenfilename(
            title="Select student Excel or CSV file",
            filetypes=[
                ("Excel and CSV files", "*.xlsx *.csv"),
                ("Excel workbook", "*.xlsx"),
                ("CSV file", "*.csv"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not filename:
        raise SystemExit("No input file selected.")

    return Path(filename)


def choose_photo_folder() -> Path:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        folder = filedialog.askdirectory(
            title="Select the folder containing student photos"
        )
    finally:
        root.destroy()

    if not folder:
        raise SystemExit("No photo folder selected.")

    return Path(folder)


# ---------------------------------------------------------
# Read Excel or CSV
# ---------------------------------------------------------

def normalize_header(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def normalize_register(value) -> str:
    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        value = int(value)

    return str(value).strip()


def normalize_dob(value) -> str:
    """Return DOB as YYYY-MM-DD for the portal login."""
    if value is None or str(value).strip() == "":
        raise ValueError("DOB is empty")

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    text = str(value).strip()

    formats = [
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d.%m.%Y",
        "%Y/%m/%d",
        "%d-%b-%Y",
        "%d %b %Y",
        "%d %B %Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    raise ValueError(
        f"Unsupported DOB format: {text!r}. "
        "Use YYYY-MM-DD or a recognizable date format."
    )


def read_student_records(file_path: Path) -> list[dict]:
    records = []

    if file_path.suffix.lower() == ".csv":
        with file_path.open(
            "r", encoding="utf-8-sig", newline=""
        ) as file:
            reader = csv.DictReader(file)

            if not reader.fieldnames:
                raise ValueError("The CSV file has no header row.")

            headers = {
                normalize_header(name): name
                for name in reader.fieldnames
            }

            reg_column = next(
                (
                    headers[key]
                    for key in (
                        "regno",
                        "registernumber",
                        "registernumber",
                        "registrationnumber",
                    )
                    if key in headers
                ),
                None,
            )

            dob_column = next(
                (
                    headers[key]
                    for key in (
                        "dob",
                        "dateofbirth",
                        "birthdate",
                    )
                    if key in headers
                ),
                None,
            )

            if not reg_column or not dob_column:
                raise ValueError(
                    "CSV must contain regno and dob columns. "
                    f"Found: {reader.fieldnames}"
                )

            for row_number, row in enumerate(reader, start=1):
                records.append({
                    "row": row_number,
                    "regno": normalize_register(row.get(reg_column)),
                    "dob_raw": row.get(dob_column),
                })

    elif file_path.suffix.lower() == ".xlsx":
        workbook = load_workbook(
            file_path,
            read_only=True,
            data_only=True,
        )

        try:
            sheet = workbook.active
            rows = sheet.iter_rows(values_only=True)
            headers_row = next(rows, None)

            if not headers_row:
                raise ValueError("The Excel sheet is empty.")

            headers = {
                normalize_header(name): index
                for index, name in enumerate(headers_row)
                if name is not None
            }

            reg_index = next(
                (
                    headers[key]
                    for key in (
                        "regno",
                        "registernumber",
                        "registrationnumber",
                    )
                    if key in headers
                ),
                None,
            )

            dob_index = next(
                (
                    headers[key]
                    for key in (
                        "dob",
                        "dateofbirth",
                        "birthdate",
                    )
                    if key in headers
                ),
                None,
            )

            if reg_index is None or dob_index is None:
                raise ValueError(
                    "Excel must contain regno and dob columns. "
                    f"Found: {list(headers_row)}"
                )

            for row_number, row in enumerate(rows, start=1):
                reg_value = (
                    row[reg_index] if reg_index < len(row) else None
                )
                dob_value = (
                    row[dob_index] if dob_index < len(row) else None
                )

                if reg_value is None and dob_value is None:
                    continue

                records.append({
                    "row": row_number,
                    "regno": normalize_register(reg_value),
                    "dob_raw": dob_value,
                })

        finally:
            workbook.close()

    else:
        raise ValueError("Select an .xlsx or .csv file.")

    if not records:
        raise ValueError("No student records found.")

    return records


# ---------------------------------------------------------
# Match photos to register numbers
# ---------------------------------------------------------

def match_key(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", value).upper()


def build_photo_index(photo_folder: Path) -> dict[str, Path]:
    index = {}

    for path in photo_folder.iterdir():
        if not path.is_file():
            continue

        if path.suffix.lower() not in IMAGE_EXTENSIONS:
            continue

        key = match_key(path.stem)

        if key in index:
            print(
                f"WARNING: Duplicate photo match for {path.stem!r}. "
                f"Using {index[key].name!r}; ignoring {path.name!r}."
            )
            continue

        index[key] = path

    return index


# ---------------------------------------------------------
# Browser and photo helpers
# ---------------------------------------------------------

def create_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1440,1100")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--headless=new")

    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_script_timeout(15)
    driver.set_page_load_timeout(45)

    return driver


def photo_data_uri(photo_path: Path) -> str:
    if not photo_path.is_file():
        raise FileNotFoundError(
            f"Photo file not found: {photo_path}"
        )

    mime = mimetypes.guess_type(photo_path.name)[0] or "image/png"
    encoded = base64.b64encode(
        photo_path.read_bytes()
    ).decode("ascii")

    return f"data:{mime};base64,{encoded}"


# ---------------------------------------------------------
# Login and open hall ticket
# ---------------------------------------------------------

def login_and_open_hallticket(
    driver: webdriver.Chrome,
    register_no: str,
    dob: str,
) -> None:
    wait = WebDriverWait(driver, WAIT_SECONDS)
    driver.get(PORTAL_URL)

    reg = wait.until(
        EC.element_to_be_clickable((By.NAME, "regno"))
    )
    reg.clear()
    reg.send_keys(register_no)

    portal_dob = datetime.strptime(
        dob, "%Y-%m-%d"
    ).strftime("%m-%d-%Y")

    dob_input = wait.until(
        EC.element_to_be_clickable((By.NAME, "dob"))
    )
    dob_input.clear()
    dob_input.send_keys(portal_dob)

    submit = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button[type='submit']")
        )
    )
    submit.click()

    wait.until(
        lambda d: "dashboard" in d.current_url.lower()
    )

    driver.get(HALLTICKET_URL)

    wait.until(
        lambda d: d.execute_script(
            "return document.readyState"
        ) == "complete"
    )

    wait.until(
        lambda d: bool(d.find_elements(By.TAG_NAME, "body"))
    )


# ---------------------------------------------------------
# Embed photo into the hall-ticket HTML
# ---------------------------------------------------------

def inject_photo(
    driver: webdriver.Chrome,
    data_uri: str,
    register_no: str,
) -> None:
    selector = f'.ht-report .ht-photo img[alt="{register_no}"]'
    matches = driver.find_elements(By.CSS_SELECTOR, selector)

    if not matches:
        matches = driver.find_elements(
            By.CSS_SELECTOR, ".ht-report td.ht-photo img"
        )

    if not matches:
        raise RuntimeError(
            "Could not locate the hall-ticket photo. "
            "Confirm the hall-ticket page has loaded."
        )

    photo = matches[0]

    driver.execute_script(
        """
        const img = arguments[0];
        const dataUri = arguments[1];

        img.removeAttribute("srcset");
        img.removeAttribute("sizes");
        img.loading = "eager";
        img.decoding = "sync";
        img.src = dataUri;

        img.style.setProperty("display", "block", "important");
        img.style.setProperty("visibility", "visible", "important");
        img.style.setProperty("opacity", "1", "important");
        img.style.setProperty("width", "1in", "important");
        img.style.setProperty("height", "1.25in", "important");
        img.style.setProperty("max-width", "none", "important");
        img.style.setProperty("object-fit", "cover", "important");
        img.style.setProperty("object-position", "center", "important");

        const cell = img.closest(".ht-photo");
        if (cell) {
            cell.style.setProperty("visibility", "visible", "important");
        }
        """,
        photo,
        data_uri,
    )

    loaded = driver.execute_async_script(
        """
        const img = arguments[0];
        const done = arguments[arguments.length - 1];

        const finish = async () => {
            try {
                if (img.decode) await img.decode();
                done(img.complete && img.naturalWidth > 0);
            } catch (e) {
                done(false);
            }
        };

        if (img.complete && img.naturalWidth > 0) {
            finish();
        } else {
            img.onload = finish;
            img.onerror = () => done(false);
        }
        """,
        photo,
    )

    if not loaded:
        raise RuntimeError("The photo failed to load or decode.")

    driver.execute_script(
        """
        const oldStyle = document.getElementById(
            "local-photo-print-fix"
        );
        if (oldStyle) oldStyle.remove();

        const style = document.createElement("style");
        style.id = "local-photo-print-fix";
        style.textContent = `
            @media print {
                .ht-report .ht-photo,
                .ht-report .ht-photo img {
                    visibility: visible !important;
                    opacity: 1 !important;
                }

                .ht-report .ht-photo img {
                    display: block !important;
                    width: 1in !important;
                    height: 1.25in !important;
                    object-fit: cover !important;
                    object-position: center !important;
                }
            }
        `;
        document.head.appendChild(style);
        """
    )

    print("Photo embedded successfully.")

# ---------------------------------------------------------
# Save the standalone HTML copy
# ---------------------------------------------------------

def save_html(
    driver: webdriver.Chrome,
    output_path: Path,
) -> None:
    driver.execute_script(
        """
        let base = document.querySelector("head base");

        if (!base) {
            base = document.createElement("base");
            document.head.prepend(base);
        }

        base.href = arguments[0];

        document.querySelectorAll('input[type="password"]')
            .forEach(el => el.value = "");
        """,
        ASSET_BASE_URL,
    )

    html = driver.execute_script(
        "return '<!DOCTYPE html>\\n' + "
        "document.documentElement.outerHTML;"
    )

    output_path.write_text(html, encoding="utf-8")


# ---------------------------------------------------------
# Generate PDF without opening the print dialog
# ---------------------------------------------------------

def generate_pdf(driver, pdf_path, data_uri, register_no):
    import base64
    import time
    from pathlib import Path
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.common.exceptions import TimeoutException, InvalidSessionIdException

    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    original_handle = driver.current_window_handle
    print_handle = None

    try:
        driver.switch_to.new_window("tab")
        print_handle = driver.current_window_handle

        driver.get("https://slmgaccoe.edecampus.com/student/HallticketPrint.php")

        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script(
                "return document.readyState"
            ) in ("interactive", "complete")
        )

        time.sleep(1)

        result = driver.execute_script("""
            const dataUri = arguments[0];
            const regNo = arguments[1];

            const imgs = Array.from(document.images);

            let img = imgs.find(i =>
                (i.alt || "").includes(regNo)
            );

            if (!img) {
                img = document.querySelector(
                    ".ht-photo img, td.ht-photo img"
                );
            }

            if (!img) {
                img = imgs.find(i =>
                    i.naturalWidth > 0 &&
                    !/logo|header|footer|icon/i.test(
                        (i.src || "") + " " + (i.alt || "")
                    )
                );
            }

            if (img) {
                img.src = dataUri;
                img.removeAttribute("srcset");
                img.style.display = "block";
                img.style.visibility = "visible";
                img.style.objectFit = "contain";
            }

            window.print = function() {};

            return {
                photoFound: !!img,
                imageCount: imgs.length,
                title: document.title
            };
        """, data_uri, str(register_no))

        print("Print page:", result["title"])
        print("Images found:", result["imageCount"])

        if not result["photoFound"]:
            print(
                "WARNING: Could not identify the photo element. "
                "The PDF may not contain the photo."
            )
        else:
            time.sleep(0.5)

        pdf_data = driver.execute_cdp_cmd(
            "Page.printToPDF",
            {
                "printBackground": True,
                "preferCSSPageSize": True,
                "displayHeaderFooter": False
            }
        )

        pdf_path.write_bytes(base64.b64decode(pdf_data["data"]))
        print(f"PDF saved automatically: {pdf_path}")

    except TimeoutException:
        raise RuntimeError(
            "The hall-ticket print page did not load in time. "
            "Check your internet connection and portal login session."
        )

    finally:
        try:
            if print_handle:
                driver.switch_to.window(print_handle)
                driver.close()
        except Exception:
            pass

        try:
            if original_handle:
                driver.switch_to.window(original_handle)
        except Exception:
            pass

# ---------------------------------------------------------
# Main bulk processing
# ---------------------------------------------------------

def generate_bulk(input_file: Path, photo_folder: Path, output_root: Path, progress_callback=None) -> dict:
    """Generate each student's HTML/PDF and save a CSV result report."""
    input_file, photo_folder, output_root = Path(input_file), Path(photo_folder), Path(output_root)
    records = read_student_records(input_file)
    photo_index = build_photo_index(photo_folder)
    html_dir, pdf_dir = output_root / "html", output_root / "pdf"
    html_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    report_rows, seen_registers = [], set()
    total = len(records)

    for position, record in enumerate(records, start=1):
        register_no, row_number = record["regno"], record["row"]
        status, message, driver = "FAILED", "", None
        try:
            if not register_no:
                raise ValueError(f"Register number is empty in spreadsheet row {row_number}.")
            reg_key = match_key(register_no)
            if reg_key in seen_registers:
                raise ValueError("Duplicate register number in the input file.")
            seen_registers.add(reg_key)
            dob = normalize_dob(record["dob_raw"])
            photo_path = photo_index.get(reg_key)
            if not photo_path:
                raise FileNotFoundError(
                    f"No matching photo found for {register_no}. Photo filename must match the register number."
                )
            data_uri = photo_data_uri(photo_path)
            safe_reg = "".join(c for c in register_no if c.isalnum() or c in "-_")
            if not safe_reg:
                raise ValueError("Register number contains no usable filename characters.")
            html_path = html_dir / f"hallticket_{safe_reg}.html"
            pdf_path = pdf_dir / f"hallticket_{safe_reg}.pdf"
            driver = create_driver()
            login_and_open_hallticket(driver, register_no, dob)
            inject_photo(driver, data_uri, register_no)
            save_html(driver, html_path)
            generate_pdf(driver, pdf_path, data_uri, register_no)
            status, message = "SUCCESS", "HTML and PDF generated successfully."
        except Exception as exc:
            message = f"{type(exc).__name__}: {exc}"
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
        row_result = {"row": row_number, "regno": register_no, "status": status, "message": message}
        report_rows.append(row_result)
        if progress_callback:
            progress_callback(position, total, row_result)

    report_path = output_root / "bulk_report.csv"
    with report_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=["row", "regno", "status", "message"])
        writer.writeheader()
        writer.writerows(report_rows)
    succeeded = sum(row["status"] == "SUCCESS" for row in report_rows)
    return {
        "total": total, "succeeded": succeeded, "failed": total - succeeded,
        "report_path": report_path, "html_dir": html_dir, "pdf_dir": pdf_dir,
        "report_rows": report_rows,
    }
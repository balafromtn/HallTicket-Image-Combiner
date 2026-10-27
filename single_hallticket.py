# single_hallticket.py
"""
Create a local hall-ticket HTML copy and a PDF with an embedded photo.
Run: python hallticket_scraper.py

Output:
    generated/html/hallticket_<register_number>.html
    generated/pdf/hallticket_<register_number>.pdf

For personal reference only. This does not modify the college portal.
"""

from __future__ import annotations

from datetime import datetime
import base64
import mimetypes
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

PORTAL_URL = "https://slmgaccoe.edecampus.com/student/index.php"
HALLTICKET_URL = "https://slmgaccoe.edecampus.com/student/Hallticket"
ASSET_BASE_URL = "https://slmgaccoe.edecampus.com/student/"
WAIT_SECONDS = 25

HERE = Path(__file__).resolve().parent


def create_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1440,1100")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--headless=new")

    service = ChromeService(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def photo_data_uri(photo_path: Path) -> str:
    if not photo_path.is_file():
        raise FileNotFoundError(f"Photo file not found: {photo_path}")

    mime = mimetypes.guess_type(photo_path.name)[0] or "image/png"
    encoded = base64.b64encode(photo_path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def login_and_open_hallticket(
    driver: webdriver.Chrome,
    register_no: str,
    dob: str,
) -> None:
    wait = WebDriverWait(driver, WAIT_SECONDS)
    driver.get(PORTAL_URL)

    reg = wait.until(EC.element_to_be_clickable((By.NAME, "regno")))
    reg.clear()
    reg.send_keys(register_no)

    # Input format: YYYY-MM-DD. Portal format: MM-DD-YYYY.
    portal_dob = datetime.strptime(dob, "%Y-%m-%d").strftime("%m-%d-%Y")
    dob_input = wait.until(EC.element_to_be_clickable((By.NAME, "dob")))
    dob_input.clear()
    dob_input.send_keys(portal_dob)

    submit = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button[type='submit']")
        )
    )
    submit.click()

    wait.until(lambda d: "dashboard" in d.current_url.lower())
    driver.get(HALLTICKET_URL)

    wait.until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    wait.until(lambda d: bool(d.find_elements(By.TAG_NAME, "body")))


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


def save_html(driver: webdriver.Chrome, output_path: Path) -> None:
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
        "return '<!DOCTYPE html>\\n' + document.documentElement.outerHTML;"
    )
    output_path.write_text(html, encoding="utf-8")


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

            let img = imgs.find(i => (i.alt || "").includes(regNo));

            if (!img) {
                img = document.querySelector(".ht-photo img, td.ht-photo img");
            }

            if (!img) {
                img = imgs.find(i =>
                    i.naturalWidth > 0 &&
                    !/logo|header|footer|icon/i.test((i.src || "") + " " + (i.alt || ""))
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
            print("WARNING: Could not identify the photo element. The PDF may not contain the photo.")
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
        raise RuntimeError("The hall-ticket print page did not load in time.")

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


def generate_single(register_no: str, dob: str, photo_path: Path, output_root: Path) -> dict:
    """Generate one student's local HTML and PDF hall-ticket files."""
    register_no = str(register_no).strip()
    dob = str(dob).strip()
    photo_path, output_root = Path(photo_path), Path(output_root)
    
    if not register_no or not dob:
        raise ValueError("Register number and DOB are required.")
    
    datetime.strptime(dob, "%Y-%m-%d")
    data_uri = photo_data_uri(photo_path)
    
    html_dir, pdf_dir = output_root / "html", output_root / "pdf"
    html_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    
    safe_reg = "".join(c for c in register_no if c.isalnum() or c in "-_")
    if not safe_reg:
        raise ValueError("Register number contains no usable filename characters.")
        
    html_path = html_dir / f"hallticket_{safe_reg}.html"
    pdf_path = pdf_dir / f"hallticket_{safe_reg}.pdf"
    
    driver = None
    try:
        driver = create_driver()
        login_and_open_hallticket(driver, register_no, dob)
        inject_photo(driver, data_uri, register_no)
        save_html(driver, html_path)
        generate_pdf(driver, pdf_path, data_uri, register_no)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass 
                
    return {"register_no": register_no, "html_path": html_path, "pdf_path": pdf_path}
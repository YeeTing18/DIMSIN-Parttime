"""
RPA – AI HAWB Search: Phase 1 + Phase 2 + Phase 3 (Document Cloud with nested iframe)

Phase 1:
- Ask user for ATA FROM/TO via dropdown date picker.
- Login, navigate to AI_HAWB_Search, apply ATA filter.

Phase 2:
- Loop through all HAWB links in the result grid (page by page).

Phase 3:
- For each HAWB:
    - Open HAWB.
    - Inside Edit_LeftTab iframe, click Document Cloud tab.
    - Switch into nested iframe id="document".
    - For each row in Document Cloud table:
        - Decide dropdown based on filename rules:
            * DIMxxxx -> HAWB
            * prefixes (II, ME, IN, ...) -> Import Customs
        - If no rule and filename is a PDF:
            * Open PDF in new tab, download it, read text with pdfplumber,
              classify (Commercial Invoice / Packing List / Dimerco Invoice),
              then set dropdown accordingly.
    - Close HAWB tab.
"""
import pyautogui
import os
import time
import traceback
import re
from datetime import datetime, date
# import getpass  # <-- no longer needed
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import ttk, messagebox
import firebase_admin
from firebase_admin import credentials, db
import pdfplumber  # pip install pdfplumber

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys  # not strictly needed now, but harmless
from selenium.webdriver import ActionChains      # same here

# -------------------------
# Timing & paths config
# -------------------------
SELENIUM_WAIT_TIMEOUT = 30       # WebDriverWait timeout
WAIT_MENU_CLICK = 1              # after clicking menu items
WAIT_AFTER_SEARCH = 12           # after clicking Search
WAIT_AFTER_HAWB_OPEN = 7         # after clicking a HAWB link
WAIT_AFTER_HAWB_CLOSE = 3        # after closing a HAWB tab
WAIT_AFTER_NEXT_PAGE = 5         # after clicking Next page
WAIT_AFTER_DOC_TAB = 10          # after clicking Document Cloud tab
WAIT_AFTER_DROPDOWN_OPEN = 1     # after opening dropdown
WAIT_AFTER_DROPDOWN_SELECT = 2.5 # after choosing a dropdown option
WAIT_AFTER_DOC_LINK_CLICK = 8    # after clicking the PDF link to let the new tab load
DOWNLOAD_WAIT_TIMEOUT = 60       # seconds to wait for PDF download

# Where Chrome will auto-save downloaded PDFs
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# -------------------------
# Config: Login & Menu
# -------------------------
LOGIN_URL = "https://dimsin.dimerco.com:8888/Default.aspx"
LOGIN_USER_ID = "wucLogin1_txtKeyword"
LOGIN_PASS_ID = "wucLogin1_txtPassword"
LOGIN_BUTTON_ID = "wucLogin1_btnLogin"

# Menu XPaths
XPATH_MENU_1 = '//*[@id="dVPMenu"]/ul/li[1]/div/a'
XPATH_MENU_2 = '//*[@id="dVPMenu"]/ul/li[1]/ul/li[1]/div/a'
XPATH_MENU_3 = '//*[@id="dVPMenu"]/ul/li[1]/ul/li[1]/ul/li[4]/div/a'

# AE_HAWB_Search iframe
AE_HAWB_IFRAME_SRC = '/V3New/AE_HAWB_Search/index?pageid=page1869'

# HAWB Edit_LeftTab iframe (SourceID changes, so use contains)
HAWB_EDIT_IFRAME_CSS = 'iframe[src*=/V3New/AE_HAWB/Edit_LeftTab"]'

# Inner Document Cloud iframe (inside Edit_LeftTab)
DOC_IFRAME_ID = "document"  # <iframe id="document"> ... #app ...

# ETD filter controls
ETD_FROM_ID = "_sfltfromETD"
ETD_TO_ID = "_sflttoETD"
SEARCH_BUTTON_ID = "btnSearch"

# Grid and navigation XPaths
ROW_XPATH_ALL = '//*[@id="AEHAWBSearchGrid"]/div[3]/table/tbody/tr'
ROW_LINK_XPATH_TEMPLATE = '//*[@id="AEHAWBSearchGrid"]/div[3]/table/tbody/tr[{row}]/td[1]/a'
NEXT_PAGE_XPATH = '//*[@id="AEHAWBSearchGrid"]/div[5]/a[3]/span'
CLOSE_HAWB_TAB_XPATH = '//*[@id="navTab"]/div[1]/div[1]/ul/li[3]/a[2]'
TOTAL_TEXT_XPATH = '//*[@id="AEHAWBSearchGrid"]/div[5]/span[2]'   # "21-40 of 58"

# Document Cloud tab (inside Edit_LeftTab iframe)
DOC_TAB_CSS = "#poDocumentCloud a"

# Document Cloud table XPaths (inside iframe#document)
DC_TABLE_ROWS_XPATH = '//*[@id="app"]/main/div/div[1]/div/div[1]/div[2]/div/table/tr[position()>1]'
# dropdown root (whole control for that row)
DC_ROW_DROPDOWN_ROOT_REL_XPATH = './/td[1]/div/div/div/div[1]'
# dropdown text area (we read current value from here)
DC_ROW_DROPDOWN_TEXT_REL_XPATH = './/td[1]/div/div/div/div[1]/div[2]'
# file name link
DC_ROW_FILENAME_REL_XPATH = './/td[2]/a'

# Download button in PDF viewer (new Chrome tab)
PDF_DOWNLOAD_BUTTON_XPATH = '//*[@id="icon"]'

DEBUG_PDF_CONTENT = False  # set to False to silence PDF content debug output

# -------------------------
# Date picker UI
# -------------------------
def ask_date_range_with_dropdown():
    root = tk.Tk()
    root.title("Select ETA date range")
    root.geometry("420x220")

    current = date.today()
    YEARS = [str(y) for y in range(2015, 2036)]
    MONTHS = [f"{m:02d}" for m in range(1, 13)]
    DAYS = [f"{d:02d}" for d in range(1, 32)]

    main_frame = ttk.Frame(root, padding=10)
    main_frame.pack(fill="both", expand=True)

    ttk.Label(main_frame, text="").grid(row=0, column=0, padx=5, pady=5)
    ttk.Label(main_frame, text="Year").grid(row=0, column=1, padx=5, pady=5)
    ttk.Label(main_frame, text="Month").grid(row=0, column=2, padx=5, pady=5)
    ttk.Label(main_frame, text="Day").grid(row=0, column=3, padx=5, pady=5)

    ttk.Label(main_frame, text="ETA FROM:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    from_year = ttk.Combobox(main_frame, values=YEARS, width=6, state="readonly")
    from_month = ttk.Combobox(main_frame, values=MONTHS, width=4, state="readonly")
    from_day = ttk.Combobox(main_frame, values=DAYS, width=4, state="readonly")
    from_year.set(str(current.year))
    from_month.set(f"{current.month:02d}")
    from_day.set(f"{current.day:02d}")
    from_year.grid(row=1, column=1, padx=5, pady=5)
    from_month.grid(row=1, column=2, padx=5, pady=5)
    from_day.grid(row=1, column=3, padx=5, pady=5)

    ttk.Label(main_frame, text="ETA TO:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
    to_year = ttk.Combobox(main_frame, values=YEARS, width=6, state="readonly")
    to_month = ttk.Combobox(main_frame, values=MONTHS, width=4, state="readonly")
    to_day = ttk.Combobox(main_frame, values=DAYS, width=4, state="readonly")
    to_year.set(str(current.year))
    to_month.set(f"{current.month:02d}")
    to_day.set(f"{current.day:02d}")
    to_year.grid(row=2, column=1, padx=5, pady=5)
    to_month.grid(row=2, column=2, padx=5, pady=5)
    to_day.grid(row=2, column=3, padx=5, pady=5)

    result = {"from": None, "to": None}

    def on_ok():
        try:
            fy, fm, fd = from_year.get(), from_month.get(), from_day.get()
            ty, tm, td = to_year.get(), to_month.get(), to_day.get()
            from_dt = datetime(int(fy), int(fm), int(fd)).date()
            to_dt = datetime(int(ty), int(tm), int(td)).date()
            if from_dt > to_dt:
                messagebox.showerror("Invalid range", "FROM date cannot be after TO date.", parent=root)
                return
            result["from"] = from_dt.strftime("%Y-%m-%d")
            result["to"] = to_dt.strftime("%Y-%m-%d")
            root.destroy()
        except ValueError:
            messagebox.showerror("Invalid date", "Please choose valid calendar dates.", parent=root)

    def on_cancel():
        result["from"] = None
        result["to"] = None
        root.destroy()

    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=3, column=0, columnspan=4, pady=15)
    ttk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=10)
    ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="left", padx=10)

    root.mainloop()

    if not result["from"] or not result["to"]:
        raise SystemExit("User cancelled date input.")

    return result["from"], result["to"]


# -------------------------
# Login popup UI
# -------------------------
def ask_login_credentials_with_popup():
    """
    Show a small Tkinter window asking for:
    - ValuePlus login ID (e.g. Y9999)
    - Password (masked)

    Returns (username, password).
    Exits the script if user cancels.
    """
    root = tk.Tk()
    root.title("ValuePlus Login")
    root.geometry("360x180")

    main_frame = ttk.Frame(root, padding=10)
    main_frame.pack(fill="both", expand=True)

    ttk.Label(main_frame, text="Please enter your ValuePlus credentials").grid(
        row=0, column=0, columnspan=2, pady=(0, 10)
    )

    ttk.Label(main_frame, text="Login ID (e.g. Y9999):").grid(
        row=1, column=0, sticky="e", padx=5, pady=5
    )
    username_var = tk.StringVar()
    username_entry = ttk.Entry(main_frame, textvariable=username_var, width=25)
    username_entry.grid(row=1, column=1, padx=5, pady=5)

    ttk.Label(main_frame, text="Password:").grid(
        row=2, column=0, sticky="e", padx=5, pady=5
    )
    password_var = tk.StringVar()
    password_entry = ttk.Entry(main_frame, textvariable=password_var, width=25, show="*")
    password_entry.grid(row=2, column=1, padx=5, pady=5)

    result = {"user": None, "pwd": None}

    def on_ok():
        user = username_var.get().strip()
        pwd = password_var.get()
        if not user or not pwd:
            messagebox.showerror("Missing data", "Please enter both login ID and password.", parent=root)
            return
        result["user"] = user
        result["pwd"] = pwd
        root.destroy()

    def on_cancel():
        result["user"] = None
        result["pwd"] = None
        root.destroy()

    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=3, column=0, columnspan=2, pady=15)
    ttk.Button(button_frame, text="OK", command=on_ok).pack(side="left", padx=10)
    ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="left", padx=10)

    # Focus username by default
    username_entry.focus_set()
    root.mainloop()

    if not result["user"] or not result["pwd"]:
        raise SystemExit("User cancelled login.")

    return result["user"], result["pwd"]


# -------------------------
# Selenium helpers
# -------------------------
def create_driver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)

    # Configure Chrome to auto-download without asking
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        # Make PDFs download instead of opening in Chrome's internal viewer
        "plugins.always_open_pdf_externally": False,
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    return driver


def login(driver, username, password):
    wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)
    driver.get(LOGIN_URL)
    user_box = wait.until(EC.presence_of_element_located((By.ID, LOGIN_USER_ID)))
    pass_box = wait.until(EC.presence_of_element_located((By.ID, LOGIN_PASS_ID)))
    user_box.clear()
    user_box.send_keys(username)
    pass_box.clear()
    pass_box.send_keys(password)
    login_btn = wait.until(EC.element_to_be_clickable((By.ID, LOGIN_BUTTON_ID)))
    login_btn.click()
    wait.until(EC.presence_of_element_located((By.XPATH, XPATH_MENU_1)))


def switch_to_ae_hawb_iframe(driver):
    driver.switch_to.default_content()
    WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT).until(
        EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, f'iframe[src="{AE_HAWB_IFRAME_SRC}"]'))
    )


def switch_to_hawb_edit_iframe(driver):
    driver.switch_to.default_content()
    # WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT).until(
    #     EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, HAWB_EDIT_IFRAME_CSS))
    # )


def switch_to_doc_cloud_iframe(driver):
    # already inside Edit_LeftTab; just go into inner iframe
    WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, DOC_IFRAME_ID))
    )


def navigate_to_ae_hawb_search(driver):
    wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)
    menu1 = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_MENU_1)))
    menu1.click()
    time.sleep(WAIT_MENU_CLICK)
    menu2 = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_MENU_2)))
    menu2.click()
    time.sleep(WAIT_MENU_CLICK)
    menu3 = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_MENU_3)))
    menu3.click()
    switch_to_ae_hawb_iframe(driver)

from datetime import datetime, timedelta

def get_last_week_range():
    today = datetime.now().date()
    
    # weekday(): Monday=0, Sunday=6
    # To get to last Sunday:
    # If today is Monday(0), we go back 1 + 7 days.
    # Formula: (today.weekday() + 1)
    days_to_last_saturday = (today.weekday() + 2) % 7
    if days_to_last_saturday == 0: # If today is Friday/Saturday, ensure we get the *previous* week
        days_to_last_saturday = 7
        
    last_saturday = today - timedelta(days=days_to_last_saturday)
    last_sunday = last_saturday - timedelta(days=6)
    
    # Format for Dimerco system (DD/MM/YYYY or your required format)
    # Adjust format string "%d/%m/%Y" if your system needs MM/DD/YYYY
    etd_from = last_sunday.strftime("%Y-%m-%d")

    etd_to = last_saturday.strftime("%Y-%m-%d")
    
    return etd_from, etd_to

def apply_etd_filter(driver, etd_from, etd_to):
    wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)
    from_input = wait.until(EC.element_to_be_clickable((By.ID, ETD_FROM_ID)))
    to_input = wait.until(EC.element_to_be_clickable((By.ID, ETD_TO_ID)))
    from_input.clear()
    from_input.send_keys(etd_from)
    to_input.clear()
    to_input.send_keys(etd_to)
    search_btn = wait.until(EC.element_to_be_clickable((By.ID, SEARCH_BUTTON_ID)))
    search_btn.click()
    #search_btn.click()

    time.sleep(WAIT_AFTER_SEARCH)


# -------------------------
# Phase 2 helpers (grid)
# -------------------------
def get_paging_info(driver):
    try:
        wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)
        total_span = wait.until(EC.presence_of_element_located((By.XPATH, TOTAL_TEXT_XPATH)))
        txt = total_span.text.strip()
        m = re.search(r'(\d+)\s*-\s*(\d+)\s*of\s*(\d+)', txt)
        if m:
            start = int(m.group(1))
            end = int(m.group(2))
            total = int(m.group(3))
            print(f"[Phase 2] Paging text: '{txt}' -> start={start}, end={end}, total={total}")
            return start, end, total
        else:
            print(f"[Phase 2] Could not parse paging text: '{txt}'")
            return None, None, None
    except Exception as e:
        print(f"[Phase 2] Failed to read paging info: {e}")
        return None, None, None


def get_row_count(driver):
    rows = driver.find_elements(By.XPATH, ROW_XPATH_ALL)
    return len(rows)


def click_hawb_link_by_row(driver, row_index):
    wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)
    hawb_xpath = ROW_LINK_XPATH_TEMPLATE.format(row=row_index)
    link = wait.until(EC.element_to_be_clickable((By.XPATH, hawb_xpath)))
    hawb_text = link.text.strip()
    link.click()
    return hawb_text


def close_hawb_tab(driver):
    driver.switch_to.default_content()
    wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)
    close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, CLOSE_HAWB_TAB_XPATH)))
    close_btn.click()
    time.sleep(WAIT_AFTER_HAWB_CLOSE)


def go_to_next_page(driver):
    wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)
    try:
        next_btn = wait.until(EC.element_to_be_clickable((By.XPATH, NEXT_PAGE_XPATH)))
        next_btn.click()
        time.sleep(WAIT_AFTER_NEXT_PAGE)
        return True
    except TimeoutException:
        return False
    except Exception:
        return False


# -------------------------
# Phase 3 – Document Cloud logic
# -------------------------
# Prefixes for "Import Customs" based on first two letters of file name
IMPORT_CUSTOMS_PREFIXES = {
    "II", "ME", "IN", "IZ", "IR", "IM", "IT", "ID", "IG",
    "OO", "OD", "OT", "TT", "TW",
}


def decide_new_doc_type(current_type_text: str, filename: str):
    """
    Decide what the dropdown SHOULD be, based on the filename only.

    Rules:
    - If filename starts with "DIM" -> target token "HAWB"
    - If first 2 letters in IMPORT_CUSTOMS_PREFIXES -> target token "Import Customs"
    - Otherwise -> None (no rule / no change here; may fall back to PDF analysis)
    """
    if not filename:
        return None

    upper_name = filename.strip().upper()
    prefix2 = upper_name[:2]

    # Rule 1: DIM... -> HAWB
    if upper_name.startswith("DIM"):
        # We look for any option whose display text contains "HAWB"
        return "HAWB"

    # Rule 2: Prefix list -> Import Customs
    if prefix2 in IMPORT_CUSTOMS_PREFIXES:
        # We look for any option whose display text contains "IMPORT CUSTOMS"
        return "Import Customs"

    # Add more rules here later as needed
    return None


def set_dc_dropdown_value(driver, row_elem, new_value: str, max_wait: int = 10):
    """
    Open the dropdown for this row and choose the option whose text CONTAINS new_value
    (case-insensitive), by clicking the visible option element via JavaScript.

    new_value here is a *token* like "HAWB", "Import Customs", "Commercial Invoice",
    "Packing List", "Dimerco Invoice", or "Others".
    """

    wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)

    # 1) Elements inside this row
    dropdown_root = row_elem.find_element(By.XPATH, DC_ROW_DROPDOWN_ROOT_REL_XPATH)
    text_elem = row_elem.find_element(By.XPATH, DC_ROW_DROPDOWN_TEXT_REL_XPATH)

    token = new_value.upper()

    # 2) Open dropdown via JS (avoids 'not interactable' on zero-size elements)
    driver.execute_script("arguments[0].click();", dropdown_root)
    time.sleep(WAIT_AFTER_DROPDOWN_OPEN)

    # 3) If it's already something with our token, just return
    try:
        current = text_elem.text.strip()
    except Exception:
        current = ""

    print(f"        [dropdown] current value before reselection: '{current}' "
          f"(target contains '{new_value}')")

    # Helper to try a given xpath strategy
    def try_click_with_xpath(xpath_description: str, xpath_expr: str) -> bool:
        """
        Return True if we successfully set the dropdown to include token, else False.
        """
        print(f"        [dropdown] trying options via {xpath_description}")

        def visible_option_present(d):
            elems = d.find_elements(By.XPATH, xpath_expr)
            return any(e.is_displayed() for e in elems)

        wait.until(visible_option_present)

        candidates = driver.find_elements(By.XPATH, xpath_expr)
        print(f"        [dropdown] {len(candidates)} candidate(s) found via {xpath_description}")

        for e in candidates:
            try:
                cand = e
                # Climb up a few levels if needed to find a clickable container
                for _ in range(5):
                    if cand.is_displayed():
                        driver.execute_script("arguments[0].click();", cand)
                        time.sleep(WAIT_AFTER_DROPDOWN_SELECT)

                        # Check what the dropdown shows now
                        try:
                            new_text = text_elem.text.strip()
                        except Exception:
                            new_text = ""

                        print(
                            f"        [dropdown candidate click - {xpath_description}] "
                            f"option='{e.text.strip()}' -> now showing: '{new_text}'"
                        )

                        if token in new_text.upper():
                            print(f"        [dropdown] target '{new_value}' CONFIRMED.")
                            return True

                        # Didn't change to what we want, try next candidate
                        break

                    cand = cand.find_element(By.XPATH, "..")

            except Exception as click_err:
                print(f"        [dropdown] error clicking candidate ({xpath_description}): {click_err}")
                continue

        return False

    # Strategy 1: look for <li> options (common for dropdown lists)
    li_xpath = (
        "//li[contains(translate(normalize-space(.), "
        "'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), "
        f"'{token}')]"
    )
    if try_click_with_xpath("LI", li_xpath):
        return

    # Strategy 2: fallback to any element containing the token (old behaviour)
    any_xpath = (
        "//*[contains(translate(normalize-space(text()), "
        "'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), "
        f"'{token}')]"
    )
    if try_click_with_xpath("ANY", any_xpath):
        return

    # If we got here, we never saw the dropdown text change to include the token
    raise Exception(f"Could not set dropdown to value containing '{new_value}'")


# -------------------------
# PDF helpers (invoice detection)
# -------------------------
def wait_for_pdf_download(download_dir: str, start_time: float, timeout: int = DOWNLOAD_WAIT_TIMEOUT):
    """
    Wait for a PDF whose modification time is AFTER start_time.
    Returns the full path, or None on timeout.
    This works even if the file name already existed before (we look at mtime).
    """
    end_time = time.time() + timeout

    while time.time() < end_time:
        time.sleep(1)

        # All PDFs currently in the folder
        pdf_paths = [
            os.path.join(download_dir, f)
            for f in os.listdir(download_dir)
            if f.lower().endswith(".pdf")
        ]
        if not pdf_paths:
            continue

        # Files modified after we started waiting (small fudge factor)
        recent = [
            p for p in pdf_paths
            if os.path.getmtime(p) >= start_time - 1
        ]
        if not recent:
            continue

        newest = max(recent, key=os.path.getmtime)

        # Make sure it's fully downloaded (no .crdownload twin)
        if not os.path.exists(newest + ".crdownload"):
            return newest

    return None


def classify_invoice_type_from_pdf(pdf_path: str) -> str | None:
    """
    Read a PDF and classify it as:
    - "Dimerco Invoice"
    - "Packing List"
    - "Commercial Invoice"
    or None if nothing matches.
    Also prints some debug info if DEBUG_PDF_CONTENT is True.

    NOTE:
    If both "INVOICE" and "PACKING LIST" appear anywhere in the text,
    we PRIORITIZE this as "Commercial Invoice".
    """
    try:
        text_chunks = []
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text() or ""
                text_chunks.append(page_text)

        full_text_raw = "\n".join(text_chunks)
        full_text = full_text_raw.upper()

        if DEBUG_PDF_CONTENT:
            print("        [PDF DEBUG] ------------------------------")
            print(f"        [PDF DEBUG] File: {pdf_path}")
            print(f"        [PDF DEBUG] Length of extracted text: {len(full_text_raw)} characters")

            # show only the first X characters so console doesn't explode
            preview_len = 600
            preview = full_text_raw[:preview_len]
            print("        [PDF DEBUG] Text preview (first "
                  f"{min(preview_len, len(full_text_raw))} chars):")
            print("        [PDF DEBUG] ---------------------------------")
            for line in preview.splitlines():
                print("        [PDF DEBUG] " + line)
            if len(full_text_raw) > preview_len:
                print("        [PDF DEBUG] ... (truncated)")
            print("        [PDF DEBUG] ------------------------------")

        # --- Keyword flags ---
        has_dimerco_invoice = "DIMERCO INVOICE" in full_text
        has_packing_list = "PACKING LIST" in full_text
        has_invoice_word = "INVOICE" in full_text  # includes COMMERCIAL INVOICE etc.

        # Most specific / strongest rules first
        if has_dimerco_invoice:
            if DEBUG_PDF_CONTENT:
                print("        [PDF DEBUG] Matched keyword: 'DIMERCO INVOICE'")
            return "Dimerco Invoice"

        # If both "INVOICE" and "PACKING LIST" appear, treat as Commercial Invoice
        if has_invoice_word and has_packing_list:
            if DEBUG_PDF_CONTENT:
                print("        [PDF DEBUG] Matched BOTH 'INVOICE' and 'PACKING LIST' "
                      "-> priority 'Commercial Invoice'")
            return "Commercial Invoice"

        # Only packing list
        if has_packing_list:
            if DEBUG_PDF_CONTENT:
                print("        [PDF DEBUG] Matched keyword: 'PACKING LIST'")
            return "Packing List"

        # Only invoice-type wording
        if has_invoice_word:
            if DEBUG_PDF_CONTENT:
                print("        [PDF DEBUG] Matched keyword: 'INVOICE' (generic/commercial)")
            return "Commercial Invoice"

        if DEBUG_PDF_CONTENT:
            print("        [PDF DEBUG] No invoice / packing list keywords found.")

    except Exception as e:
        print(f"        Error reading PDF '{pdf_path}': {e}")

    return None
import pyperclip
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

def analyze_pdf_for_row(driver, row_idx, filename_elem, hawb_text):
    original_window = driver.current_window_handle
    existing_handles = set(driver.window_handles)
    
    filename = filename_elem.text.strip()
    upper_name = filename.upper()
    target_token = "Others" 

    # --- 1. Filename Rules (Fastest) ---
    #if upper_name.startswith("DIM"):
      #  print(f"      Row {row_idx}: Identified as Dimerco Invoice by filename.")
       # return "Dimerco Invoice"
    if upper_name.startswith("HAWB") and len(upper_name) >= 7 and upper_name[4:7].isdigit():
        print(f"      Row {row_idx}: Identified as HAWB by filename.")
        return "HAWB"

    # --- 2. Content Analysis (Download and Read) ---
    print(f"      Row {row_idx}: Downloading PDF for content check...")
    click_time = time.time()
    filename_elem.click() # This triggers download based on your 'create_driver' settings
    
    # Wait for the file to appear in your DOWNLOAD_DIR
    pdf_path = wait_for_pdf_download(DOWNLOAD_DIR, click_time)
    
    if pdf_path:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Extract text from all pages and combine
                full_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text.upper() + " "

            # --- Logic Hierarchy ---
            if "COMMERCIAL INVOICE" in full_text:
                target_token = "Commercial Invoice"
            elif "PACKING LIST" in full_text:
                target_token = "Packing List"
            elif "PERMIT" in full_text:
                target_token = "Export Permit"
            elif "DIMERCO INVOICE" in full_text:
                target_token = "Dimerco Invoice"
            elif "INVOICE" in full_text:
                target_token = "Commercial Invoice"
            
            print(f"      Row {row_idx}: Content analysis found -> {target_token}")

        except Exception as e:
            print(f"      Row {row_idx}: Failed to read PDF: {e}")
        finally:
            # Clean up: delete the file after reading so the folder stays clean
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
    else:
        print(f"      Row {row_idx}: Download failed or timed out.")

    return target_token
def classify_permit_via_screenshot(driver, pdf_path_for_reference):
    """
    Captures a screenshot of the current tab (PDF viewer) 
    and checks if it should be classified as an Export Permit.
    """
    screenshot_path = pdf_path_for_reference.replace(".pdf", ".png")
    driver.save_screenshot(screenshot_path)
    print(f"        [Screenshot] Saved to {screenshot_path}")

    # Note: To 'read' the screenshot, you'd typically use pytesseract.
    # If you want to stick to the requirement of "finding the word permit":
    # We will use the existing text extraction as a fallback or primary check
    # since 'Permit' is a text-based search.
    
    text = ""
    with pdfplumber.open(pdf_path_for_reference) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "").upper()

    if "PERMIT" in text:
        print("        [Classification] 'PERMIT' word found -> Export Permit")
        return "Export Permit"
    
    return None
import pyperclip
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd # Add this at the top of your script

# Global list to store data for Excel
excel_results = []

def append_to_excel(hawb_no, status, pdf_name, old_type, new_type):
    """Adds a new row of data to our global list."""
    excel_results.append({
        "HAWB No": hawb_no,
        "Status": status,
        "PDF Name": pdf_name,
        "Old Type": old_type,
        "New Type": new_type
    })
    
    # Save to Excel immediately so data isn't lost if script crashes
    df = pd.DataFrame(excel_results)
    df.to_excel("AE_DocCloud.xlsx", index=False)

def process_document_cloud_for_current_hawb(driver, hawb_text=None):
    changed_count = 0
    
    # 1. Initial Switch into the nested iframes
    switch_to_hawb_edit_iframe(driver)
    wait = WebDriverWait(driver, 20)

    try:
        # Switch to the tab iframe first
        iframe_tab = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='/V3New/AE_HAWB/Edit_LeftTab']")))
        driver.switch_to.frame(iframe_tab)
        
        # Click Document Cloud tab
        doc_cloud_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Document Cloud']")))
        doc_cloud_btn.click()
        time.sleep(2)
        
        # Switch to the inner 'document' iframe
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "document")))
    except Exception as e:
        print(f"    [{hawb_text}] Could not access Document Cloud: {e}")
        return 0

    # 2. Get total rows
    try:
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, DC_TABLE_ROWS_XPATH)))
        row_count = len(rows)
    except:
        return 0

    print(f"    [{hawb_text}] Found {row_count} rows. Starting verification...")

    for idx in range(1, row_count + 1):
        try:
            # Re-locate the row to avoid stale elements
            current_row_xpath = f"({DC_TABLE_ROWS_XPATH})[{idx}]"
            row = driver.find_element(By.XPATH, current_row_xpath)
            
            # Read Current Filename
            filename_elem = row.find_element(By.XPATH, DC_ROW_FILENAME_REL_XPATH)
            current_type = row.find_element(By.XPATH, DC_ROW_DROPDOWN_TEXT_REL_XPATH).text.strip()
            filename = filename_elem.text.strip()
            upper_name = filename.upper()

            # READ ORIGINAL TYPE (To check if we need to change it)
            try:
                current_type_elem = row.find_element(By.XPATH, DC_ROW_DROPDOWN_TEXT_REL_XPATH)
                original_type = current_type_elem.text.strip()
            except:
                original_type = ""
            
            final_type = None

            # --- Rule A: Filename Rules ---
            #if upper_name.startswith("DIM"):
            #    final_type = "Dimerco Invoice"
            name_without_ext = upper_name.replace(".PDF", "").strip()
            
            if hawb_text and name_without_ext == hawb_text.upper():
                final_type = "HAWB"
                print(f"      Row {idx}: {filename} is an EXACT match for HAWB -> HAWB")

            # --- RULE 2: Starts with HAWB + 3 digits ---
            # e.g. HAWB123...
            elif hawb_text and len(hawb_text) >= 6 and upper_name[:3] == hawb_text[3:6]:
                print(f"      Row {idx}: Partial match for HAWB ({upper_name[:6]}) -> HAWB")
                final_type = "Dimerco Invoice"
                print(f"      Row {idx}: Match found ({upper_name[:3]}) -> Dimerco Invoice")
            # elif upper_name.endswith(".JPG") or upper_name.endswith(".PNG"):
            #     final_type = "Others"
            
         # --- Rule B: PDF Content Analysis (Improved Focus) ---
            elif upper_name.endswith(".PDF"):
                original_window = driver.current_window_handle
                existing_handles = set(driver.window_handles)
                
                filename_elem.click()
                wait.until(lambda d: len(d.window_handles) > len(existing_handles))
                new_tab = (set(driver.window_handles) - existing_handles).pop()
                driver.switch_to.window(new_tab)
                
                try:
                    # Clear clipboard first to ensure we aren't reading old data
                    pyperclip.copy("") 
                   
                    # Wait for PDF to render
                    time.sleep(6) 

                    # Try to focus the PDF content specifically
                    actions = ActionChains(driver)
                    # Instead of clicking 'body', we send a TAB key to move focus into the PDF plugin
                    actions.send_keys(Keys.TAB).perform() 
                    time.sleep(0.5)
                    actions.send_keys(Keys.TAB).perform() 
                    time.sleep(0.5)
                    for _ in range(5):  # Change 5 to 10 if documents are very long
                        pyautogui.press('end')
                        time.sleep(0.3) # Small gap between presses
                    
                    time.sleep(1.5) # Final wait to let the text render
                    # Select All and Copy
                    actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                    time.sleep(2)
                    actions.key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()
                    actions.send_keys(Keys.CONTROL, 'c').perform()
                    time.sleep(1.0)
                    pyautogui.click(pyautogui.size().width / 2, pyautogui.size().height / 2)
                    time.sleep(0.5)
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.5)
                    pyautogui.hotkey('ctrl', 'c')
                    content = pyperclip.paste().upper().strip()
                    print("copy using action keys")
                    ocr_text = pyperclip.paste()
                    print("copy using pyperclip")
                    content = content.upper()
                    
                    # DEBUG: Print the first 100 characters so you can see if it worked
                    if not content:
                        print(f"      [Warning] Row {idx}: Clipboard is EMPTY. PDF text selection failed.")
                        driver.close()
                        driver.switch_to.window(original_window)
                        # Standard re-entry into iframes
                        switch_to_hawb_edit_iframe(driver)
                        driver.switch_to.frame(wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='/V3New/AE_HAWB/Edit_LeftTab']"))))
                        driver.switch_to.frame("document")
                        continue # Skip to the next row in the loop
                    else:
                        print(f"      [Debug] Row {idx}: Copied {len(content)} characters. Preview: {content[:100]}...")

                    # Decision Logic (Matches your requirements)
                    if "PERMIT NO" in content: 
                        final_type = "Export Customs"
                    elif "CONTRACT OF PURCHASE AND SALE" in content: 
                        final_type = "Commercial Invoice"
                    elif "DECLARATION" in content: 
                        final_type = "Others"
                    elif "COMMERCIAL INVOICE" in content: 
                        final_type = "Commercial Invoice"
                    elif "TAX INVOICE" in content: 
                        final_type = "Commercial Invoice"
                    elif "PACKING LIST" in content: 
                        final_type = "Packing List"
                    elif "PACKING SLIP" in content: 
                        final_type = "Packing List"
                    elif "PACKING" in content: 
                        final_type = "Packing List"
                    elif "PACKLIST" in content: 
                        final_type = "Packing List"
                    elif "INVOICE" in content: 
                        final_type = "Commercial Invoice"
                    
                    elif "DIMERCO INVOICE" in content: 
                        final_type = "Dimerco Invoice"
                    
                    # NEW SKIP LOGIC:
                    if final_type is None:
                        print(f"      [Skip] Row {idx}: No keywords found in PDF. Skipping this row.")
                        driver.close()
                        driver.switch_to.window(original_window)
                        # Standard re-entry into iframes
                        switch_to_hawb_edit_iframe(driver)
                        driver.switch_to.frame(wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='/V3New/AE_HAWB/Edit_LeftTab']"))))
                        driver.switch_to.frame("document")
                        continue # This jumps to the next 'idx' in the row loop
                    
                except Exception as focus_err:
                    print(f"      Row {idx}: Copy error: {focus_err}")
                # Close and RE-RE-ENTER Iframes
                driver.close()
                driver.switch_to.window(original_window)
                
                # Re-switch to iframes
                switch_to_hawb_edit_iframe(driver)
                iframe_tab = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='/V3New/AE_HAWB/Edit_LeftTab']")))
                driver.switch_to.frame(iframe_tab)
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "document")))
                
                row = driver.find_element(By.XPATH, current_row_xpath)

            # --- 3. COMPARE AND UPDATE ---
            # Remove '*' from original_type if it exists (e.g. '*Commercial Invoice' -> 'Commercial Invoice')
            clean_original = original_type.replace("*", "").strip()

            if final_type.upper() == clean_original.upper():
                print(f"      Row {idx}: {filename} -> Type is already correct ({original_type}). Skipping update.")
                #append_to_excel(hawb_text, "Done", filename, current_type, "No Change")
            else:
                print(f"      Row {idx}: {filename} -> Identified as {final_type}. Updating from {original_type}...")
                set_dc_dropdown_value(driver, row, final_type)
                changed_count += 1
                append_to_excel(hawb_text, "Done", filename, current_type, final_type)
        except Exception as e:
            print(f"      Row {idx} Error: {e}")
            # append_to_excel(hawb_text, "Error", filename, "N/A", str(e))
            # Reset iframes for next row attempt
            try:
                driver.switch_to.default_content()
                switch_to_hawb_edit_iframe(driver)
                driver.switch_to.frame(wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='/V3New/AE_HAWB/Edit_LeftTab']"))))
                driver.switch_to.frame("document")
            except: pass
            continue

    return changed_count
# -------------------------
# Combined Phase 2 + 3 loop
# -------------------------
import win32com.client as win32
import os
from datetime import datetime
def send_report_via_outlook_app(file_path, recipient_email):
    try:
        # Ensure path is absolute for the EXE environment
        abs_report_path = os.path.abspath(file_path)
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = recipient_email
        mail.Subject = f"AE Document Cloud Sorting Report"
        mail.Body = "Please find the attached RPA report."
        
        if os.path.exists(abs_report_path):
            mail.Attachments.Add(abs_report_path)
            mail.Send()
            print("✅ EXE successfully triggered Outlook email.")
        else:
            print(f"❌ EXE could not find file at: {abs_report_path}")
    except Exception as e:
        print(f"❌ EXE Email Error: {e}")
def phase2_and_3_loop_hawbs(driver):
    page_index = 1
    total_hawbs_opened = 0
    total_rows_changed = 0

    while True:
        switch_to_ae_hawb_iframe(driver)
        start_idx, end_idx, total_items = get_paging_info(driver)
        row_count = get_row_count(driver)
        print(f"[Phase 2] Page {page_index}: row_count={row_count}, paging={start_idx}-{end_idx} of {total_items}")

        if row_count == 0:
            print(f"[Phase 2] Page {page_index}: no rows found, stopping.")
            break

        for row_idx in range(1, row_count + 1):
            switch_to_ae_hawb_iframe(driver)
            try:
                hawb_text = click_hawb_link_by_row(driver, row_idx)
                total_hawbs_opened += 1
                print(f"  Page {page_index} - Row {row_idx}: opened HAWB '{hawb_text}'")
            except TimeoutException:
                print(f"  Page {page_index} - Row {row_idx}: no clickable link (skipping).")
                continue
            except Exception as e:
                print(f"  Page {page_index} - Row {row_idx}: error clicking HAWB link: {e}")
                continue

            time.sleep(WAIT_AFTER_HAWB_OPEN)

            try:
                changed = process_document_cloud_for_current_hawb(driver, hawb_text)
                total_rows_changed += changed
                print(f"    [{hawb_text}] Document Cloud rows changed: {changed}")
            except Exception as e:
                print(f"    [{hawb_text}] Error in Document Cloud processing: {e}")
                traceback.print_exc()

            try:
                close_hawb_tab(driver)
                print(f"    Closed HAWB tab for '{hawb_text}'.")
            except Exception as e:
                print(f"    Error closing HAWB tab for '{hawb_text}': {e}")
                time.sleep(1.0)

        if end_idx is not None and total_items is not None and end_idx >= total_items:
            print(f"[Phase 2] Last page reached (end index {end_idx} == total {total_items}). Stopping.")
            break

        switch_to_ae_hawb_iframe(driver)
        print(f"[Phase 2] Finished Page {page_index}, attempting to go to next page...")
        go_to_next_page(driver)
        if not go_to_next_page(driver):
            print("[Phase 2] Next page not clickable. Stopping.")
            break

        page_index += 1

    print(f"[Phase 2+3] Completed. Total HAWBs opened: {total_hawbs_opened}, "
          f"total rows changed: {total_rows_changed}")

def get_password_from_firebase():
    """
    Initializes Firebase Admin and retrieves the password from the Realtime Database.
    """
    try:
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        cred_path = os.path.join(desktop_path, 'serviceAccountKey.json')
        cred = credentials.Certificate(cred_path)

        if not firebase_admin._apps:
            firebase_admin.initialize_app(cred, {
                'databaseURL': 'https://y9999-2dcbd-default-rtdb.asia-southeast1.firebasedatabase.app'
            })

        ref = db.reference('users/Y9999/password')
        password = ref.get()
        if not password:
            print("❌ Firebase returned an empty password.")
            return None
        return password
    except Exception as e:
        print(f"❌ Failed to retrieve password from Firebase: {e}")
        return None
# -------------------------
# Main
# -------------------------
def main():
    try:
        # Phase 1a: ask for login credentials via popup
       # username, password = ask_login_credentials_with_popup()
        report_name = "AE_DocCloud.xlsx"
        target_email = "avry_a_how@dimerco.com;sam_s_kong@dimerco.com;may_m_lui@dimerco.com;finance_dimsin@dimerco.com"
        # Phase 1b: ask for ETD range
        etd_from, etd_to = get_last_week_range()
        print(f"Using ETD date range: {etd_from} -> {etd_to}")
        #wait = WebDriverWait(driver, SELENIUM_WAIT_TIMEOUT)
        driver = create_driver()
        driver.get(LOGIN_URL)

        print("Logging in...")
        print("🔐 Login page detected. Authenticating...")
        wait = WebDriverWait(driver, 10) # 10 is the timeout in seconds
        wait.until(EC.visibility_of_element_located((By.ID, "wucLogin1_txtKeyword"))).send_keys('Y9999')
        dimerco_password = get_password_from_firebase()
        if not dimerco_password:
            print("Could not retrieve password. Exiting.")
            return
        driver.find_element(By.ID, "wucLogin1_txtPassword").send_keys(dimerco_password)
        driver.find_element(By.ID, "wucLogin1_btnLogin").click()
        #os.wait.until(EC.url_contains("Default.aspx"))
        print("✅ Login successful.")
        #print(f"Using login ID: {username}")
        #login(driver, username, password)
        print("Login success.")

        print("Navigating to AE_HAWB_Search...")
        navigate_to_ae_hawb_search(driver)
        print("Inside AE_HAWB_Search iframe.")
        time.sleep(5)
        print("Applying ETD filter and clicking Search...")
        time.sleep(5)
        apply_etd_filter(driver, etd_from, etd_to)
        time.sleep(4)
        print("Search completed, results should be visible.")

        print("\n[Phase 2+3] Starting HAWB loop with Document Cloud processing...")
        phase2_and_3_loop_hawbs(driver)
        print("\nAll phases completed.")

        print("\nYou can now inspect the browser and console output.")
    except SystemExit as e:
        print(str(e))
    except Exception as e:
        print("ERROR during execution:")
        print(e)
        traceback.print_exc()
    finally:
        # This code runs no matter what
        if os.path.exists(report_name):
            print("Initiating final report email...")
            send_report_via_outlook_app(report_name, target_email)
        else:
            print("No report file found to email.")
        
        # Close browser
        try:
            driver.quit()
        except:
            pass


if __name__ == "__main__":
    main()

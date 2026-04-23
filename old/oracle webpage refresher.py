"""
Oracle Cloud – Work Definition Refresh Automation
==================================================
Place this script and your items.xlsx in the same folder, then run from VSCode.
"""

# ============================================================
#  USER CONFIG  ← only section you should ever need to edit
# ============================================================

EXCEL_FILENAME = "items.xlsx"
EXCEL_SHEET    = 0
EXCEL_COLUMN   = "A"
FIREFOX_BINARY = r"C:\Users\kmageshkumar\AppData\Local\Mozilla Firefox\firefox.exe"
HEADLESS       = False

# ============================================================
#  CONSTANTS
# ============================================================

ORACLE_URL  = ("https://fa-eovh-saasfaprod1.fa.ocs.oraclecloud.com"
               "/fscmUI/faces/AtkHomePageWelcome")
MAX_RETRIES = 3
POLL        = 0.4
SHORT_WAIT  = 15
MEDIUM_WAIT = 30
LONG_WAIT   = 120
LOGIN_WAIT  = 300

# ============================================================
#  IMPORTS
# ============================================================

import logging, sys, time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    StaleElementReferenceException, TimeoutException, WebDriverException)
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager

# ============================================================
#  LOGGING
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("wd_refresh.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ============================================================
#  RUN SUMMARY
# ============================================================

@dataclass
class RunSummary:
    success: list = field(default_factory=list)
    skipped: list = field(default_factory=list)
    failed:  list = field(default_factory=list)

    def report(self):
        log.info("=" * 60)
        log.info("RUN SUMMARY")
        log.info("  Successful : %d", len(self.success))
        log.info("  Skipped    : %d", len(self.skipped))
        log.info("  Failed     : %d", len(self.failed))
        if self.skipped: log.info("  Skipped : %s", self.skipped)
        if self.failed:  log.info("  Failed  : %s", self.failed)
        log.info("=" * 60)

# ============================================================
#  DRIVER
# ============================================================

def build_driver():
    opts = FirefoxOptions()
    opts.binary_location = FIREFOX_BINARY
    if HEADLESS:
        opts.add_argument("--headless")
    opts.set_preference("devtools.console.stdout.content", False)
    svc = FirefoxService(executable_path=GeckoDriverManager().install(),
                         log_path="geckodriver.log")
    driver = webdriver.Firefox(service=svc, options=opts)
    driver.maximize_window()
    return driver

# ============================================================
#  GENERIC HELPERS
# ============================================================

def W(driver, timeout):
    return WebDriverWait(driver, timeout, poll_frequency=POLL)

def find(driver, by, loc, timeout=MEDIUM_WAIT):
    return W(driver, timeout).until(EC.visibility_of_element_located((by, loc)))

def clickable(driver, by, loc, timeout=MEDIUM_WAIT):
    return W(driver, timeout).until(EC.element_to_be_clickable((by, loc)))

def maybe(driver, by, loc, timeout=5):
    try:
        return W(driver, timeout).until(EC.presence_of_element_located((by, loc)))
    except TimeoutException:
        return None

def js_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center',inline:'center'});", el)
    driver.execute_script("arguments[0].click();", el)

def js_double_click(driver, el):
    driver.execute_script("""
        var evt = new MouseEvent('dblclick', {
            bubbles: true,
            cancelable: true,
            view: window
        });
        arguments[0].dispatchEvent(evt);
    """, el)

def wait_gone(driver, by, loc, timeout=LONG_WAIT):
    W(driver, timeout).until(EC.invisibility_of_element_located((by, loc)))

def wait_spinner_gone(driver, timeout=LONG_WAIT):
    selectors = [
        (By.CSS_SELECTOR, "div.AFBusyIndicator"),
        (By.CSS_SELECTOR, "div[class*='AFBlockingGlassPane']"),
        (By.CSS_SELECTOR, "div[class*='af_loadingIndicator']"),
        (By.XPATH, "//*[contains(@class,'AFBusyIndicator')]"),
    ]
    deadline = time.time() + timeout
    while time.time() < deadline:
        visible = False
        for by, sel in selectors:
            try:
                if any(e.is_displayed() for e in driver.find_elements(by, sel)):
                    visible = True
                    break
            except (StaleElementReferenceException, WebDriverException):
                pass
        if not visible:
            return
        time.sleep(POLL)
    log.warning("Spinner still visible after %.0fs – proceeding.", timeout)

# ============================================================
#  EXCEL LOADER
# ============================================================

# ============================================================
#  STEP 8+9+10 – Delete existing Items via Context Menu
# ============================================================

def delete_items_if_present(driver):
    log.info("Checking for Items sub-box…")

    items_box = maybe(driver, By.XPATH,
        "//*[normalize-space(text())='Items']"
        "[ancestor::*[contains(translate(.,'K','k'),'kitting')]]",
        timeout=5
    )

    if items_box is None:
        log.info("No Items box – skipping deletion.")
        return

    log.info("Items box found – initiating context menu deletion.")

    # 1. Click the 'Tick' icon on the top toolbar
    log.info("Clicking the top toolbar 'Tick' icon...")
    tick_icon = find(driver, By.XPATH, 
        "//a[@title='Select'] | //button[@title='Select'] | "
        "//img[contains(@src, 'select') or contains(@src, 'check')]/.. | "
        "//*[contains(@class,'Toolbar')]//*[contains(@class,'select') or contains(@class,'check')]"
    )
    js_click(driver, tick_icon)
    time.sleep(1)

    # 2. Click the circle box on the top right of "Items"
    log.info("Clicking the selection circle on the 'Items' box...")
    items_circle = find(driver, By.XPATH,
        "//*[normalize-space(text())='Items']/ancestor::*[name()='g'][1]//*[name()='circle']"
    )
    js_click(driver, items_circle)
    time.sleep(1)

    # 3. Right-click the basket inside "Items"
    log.info("Right-clicking the Items basket icon...")
    items_basket = find(driver, By.XPATH,
        "//*[normalize-space(text())='Items']/ancestor::*[name()='g'][1]//*[name()='image' or contains(@class,'icon')]"
    )
    ActionChains(driver).context_click(items_basket).perform()
    time.sleep(1)

    # 4. Click Actions -> Delete
    actions_menu = find(driver, By.XPATH, "//*[normalize-space(text())='Actions']")
    js_click(driver, actions_menu)
    time.sleep(0.5)

    delete_menu = find(driver, By.XPATH, "//*[normalize-space(text())='Delete' and not(ancestor::button)]")
    js_click(driver, delete_menu)
    time.sleep(1)

    # 5. Confirm delete dialog
    ok = clickable(driver, By.XPATH,
        "//div[@role='dialog']//button[normalize-space()='OK']"
        " | //button[normalize-space()='OK']",
        timeout=SHORT_WAIT
    )
    js_click(driver, ok)
    log.info("Delete dialog confirmed.")

    # Wait for Items box to disappear
    wait_gone(driver, By.XPATH,
        "//*[normalize-space(text())='Items']"
        "[ancestor::*[contains(translate(.,'K','k'),'kitting')]]",
        timeout=LONG_WAIT
    )
    wait_spinner_gone(driver)
    log.info("Items deleted.")

# ============================================================
#  STEP 14+15 – Assign Basket via Context Menu
# ============================================================

def assign_items_to_kitting(driver):
    log.info("Assigning collected items via context menu…")

    # 1. Right-click the basket below item structure
    basket = find(driver, By.XPATH,
        "//*[contains(@id,'itemStructure') or contains(@id,'ItemStructure') or contains(@id,'bomStructure')]"
        "//*[contains(@class,'node') or contains(@class,'row') or contains(@class,'item')][1]"
        " | //*[contains(@class,'oj-dvt') or contains(@class,'dvt-node')][1]"
    )
    ActionChains(driver).context_click(basket).perform()
    time.sleep(1)

    # 2. Click Actions -> Assign
    actions_menu = find(driver, By.XPATH, "//*[normalize-space(text())='Actions']")
    js_click(driver, actions_menu)
    time.sleep(0.5)

    assign_menu = find(driver, By.XPATH, "//*[normalize-space(text())='Assign']")
    js_click(driver, assign_menu)
    time.sleep(1)

    # 3. Click OK in the popup window
    log.info("Confirming assignment popup...")
    ok = clickable(driver, By.XPATH,
        "//div[@role='dialog']//button[normalize-space()='OK']"
        " | //button[normalize-space()='OK']",
        timeout=SHORT_WAIT
    )
    js_click(driver, ok)
    wait_spinner_gone(driver, timeout=LONG_WAIT)

    # Confirm Items sub-box reappeared in Kitting
    try:
        W(driver, LONG_WAIT).until(EC.presence_of_element_located((
            By.XPATH,
            "//*[normalize-space(text())='Items']"
            "[ancestor::*[contains(translate(.,'K','k'),'kitting')]]"
        )))
        log.info("Assignment confirmed – Items visible in Kitting.")
    except TimeoutException:
        log.warning("Items box did not reappear – proceeding anyway.")

# ============================================================
#  PER-ITEM ORCHESTRATION (Updated)
# ============================================================

def process_item(driver, item_number) -> str:
    log.info("─" * 50)
    log.info("Processing: %s", item_number)
    search_for_item(driver, item_number)
    if not click_main_link(driver, item_number):
        return "skipped"
    
    W(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(.,'Kitting')]"))
    )
    time.sleep(2) 

    switch_to_main_frame(driver)
    double_click_kitting(driver)

    W(driver, 30).until(
        EC.presence_of_element_located((
            By.XPATH,
            "//*[name()='text' and contains(.,'Items')]"
        ))
    )
    time.sleep(1)

    # Use the new context menu flow
    delete_items_if_present(driver)
    collect_all_direct_children(driver)
    assign_items_to_kitting(driver)  # Replaces drag_basket_to_kitting
    save_and_close(driver)
    reset_search(driver)
    
    log.info("✓  %s complete.", item_number)
    return "success"

def load_items():
    script_dir = Path(__file__).parent.resolve()
    path = script_dir / EXCEL_FILENAME
    if not path.exists():
        cands = list(script_dir.glob("*.xlsx")) + list(script_dir.glob("*.xls")) + list(script_dir.glob("*.csv"))
        raise FileNotFoundError(
            f"\nCannot find '{EXCEL_FILENAME}' in:\n  {script_dir}\n"
            f"Files found: {[f.name for f in cands]}\n"
            f"Rename your file to '{EXCEL_FILENAME}' or update EXCEL_FILENAME at the top."
        )
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        df = pd.read_excel(path, sheet_name=EXCEL_SHEET, header=None, dtype=str)
    else:
        df = pd.read_csv(path, header=None, dtype=str)
    col = ord(EXCEL_COLUMN.upper()) - ord("A") if isinstance(EXCEL_COLUMN, str) else int(EXCEL_COLUMN)
    items = df.iloc[:, col].dropna().str.strip().replace("", pd.NA).dropna().unique().tolist()
    log.info("Loaded %d item(s) from '%s'.", len(items), path.name)
    return items

# ============================================================
#  STEP 1 – Navigate and login
# ============================================================

def navigate_and_login(driver):
    log.info("Opening Oracle Cloud…")
    driver.get(ORACLE_URL)
    wait_spinner_gone(driver, timeout=30)
    if maybe(driver, By.CSS_SELECTOR, "input[type='password']", timeout=5):
        log.warning("\n" + "="*55 +
            "\nLogin page detected – please log in manually in the browser." +
            "\nScript will continue automatically once logged in." +
            "\nYou have %d seconds.\n" % LOGIN_WAIT + "="*55)
        W(driver, LOGIN_WAIT).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
        wait_spinner_gone(driver, timeout=30)
        log.info("Login complete.")
    else:
        log.info("Already logged in.")

# ============================================================
#  STEP 2 – Open Manage Work Definitions via Favourites star
# ============================================================

def open_manage_work_definitions(driver):
    """
    Slide 2: Click the star icon in the top nav → click 'Manage Work Definitions'.
    The star icon in Oracle's nav bar has a specific structure we target robustly.
    """
    log.info("Opening Manage Work Definitions…")

    # Star / Favorites icon – try multiple known Oracle ADF nav bar selectors
    star = clickable(driver, By.XPATH,
        # Oracle ADF top nav bar star button – common id pattern
        "//*[contains(@id,'pt1') and contains(@id,'star')]"
        " | //*[@title='Favorites and Recent Items']"
        " | //*[contains(@class,'oj-ux-ico-star')]/.."
        " | //img[contains(@src,'star') or contains(@src,'favorite')]/.."
    )
    js_click(driver, star)
    wait_spinner_gone(driver, timeout=10)

    mwd = clickable(driver, By.XPATH,
        "//a[normalize-space()='Manage Work Definitions']"
        " | //span[normalize-space()='Manage Work Definitions']"
    )
    js_click(driver, mwd)
    wait_spinner_gone(driver)

    # Wait for the Manage Work Definitions page to confirm
    find(driver, By.XPATH,
        "//h1[normalize-space()='Manage Work Definitions']"
        " | //*[contains(@id,'ManageWorkDef')]"
        " | //span[normalize-space()='Manage Work Definitions']"
    )
    log.info("Manage Work Definitions loaded.")

# ============================================================
#  STEP 3+4 – Type item number and click Search
# ============================================================

def search_for_item(driver, item_number):
    """
    Slide 3:
      - Type item number into the Item LOV field
      - Click the Search button (id ends with '::search')
    """
    log.info("Searching for: %s", item_number)
    wait_spinner_gone(driver, timeout=15)

    # Item input field – the LOV text input next to the "Item" label
    item_input = find(driver, By.XPATH,
        "//label[normalize-space()='Item']/following::input[@type='text'][1]"
        " | //span[normalize-space()='Item']/following::input[@type='text'][1]"
        " | //td[normalize-space()='Item']/following-sibling::td[1]//input[1]"
    )
    item_input.click()
    time.sleep(0.3)
    item_input.send_keys(Keys.CONTROL + "a")
    item_input.send_keys(Keys.DELETE)
    item_input.clear()
    item_input.send_keys(item_number)
    log.info("Item number entered.")
    time.sleep(0.3)

    # Search button – confirmed via DOM inspection: id ends with '::search'
    # Using XPath substring() function which works in all browsers
    search_btn = clickable(driver, By.XPATH,
        "//button[substring(@id, string-length(@id) - 7) = '::search']"
        " | //button[normalize-space()='Search']"
        " | //input[@value='Search']"
    )
    js_click(driver, search_btn)
    log.info("Search clicked.")
    wait_spinner_gone(driver)

# ============================================================
#  STEP 5+6 – Select best result row and click Main
# ============================================================

def click_main_link(driver, item_number) -> bool:
    """
    Slide 4:
      Wait for the 'Main' link to appear in the results table.
      Pick the row where Production Priority=1 and Version is highest.
      Click its 'Main' link.
    """
    log.info("Waiting for results…")
    try:
        W(driver, MEDIUM_WAIT).until(
            EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Main']")))
    except TimeoutException:
        log.warning("No results found for %s.", item_number)
        return False

    rows = driver.find_elements(By.XPATH, "//tr[.//a[normalize-space()='Main']]")
    if not rows:
        log.warning("No rows with 'Main' link for %s.", item_number)
        return False

    def digits(s):
        d = "".join(c for c in s if c.isdigit())
        return int(d) if d else -1

    best_row, best_v = None, -1

    for row in rows:
        try:
            cells = row.find_elements(By.TAG_NAME, "td")
            texts = [c.text.strip() for c in cells]
            log.debug("Row cells: %s", texts)

            # Columns (0-based):
            # 0=Item  1=ItemDesc  2=StructureName  3=Name(Main)
            # 4=Type  5=Version   6=ProdPriority   7=CostPriority ...
            version_text = texts[5] if len(texts) > 5 else ""
            pp_text      = texts[6] if len(texts) > 6 else ""

            # Override with header-attribute cells if Oracle provides them
            vc  = row.find_elements(By.XPATH, ".//td[contains(@headers,'version') or contains(@headers,'Version')]")
            ppc = row.find_elements(By.XPATH, ".//td[contains(@headers,'ProductionPriority') or contains(@headers,'productionPriority') or contains(@headers,'production')]")
            if vc:  version_text = vc[0].text.strip()
            if ppc: pp_text      = ppc[0].text.strip()

            v  = digits(version_text)
            pp = digits(pp_text)
            log.debug("  -> version=%d  prod_priority=%d", v, pp)

            if pp == 1 and v > best_v:
                best_v   = v
                best_row = row

        except (IndexError, StaleElementReferenceException):
            continue

    if best_row is None:
        log.warning("No row with ProdPriority=1 for %s. Skipping.", item_number)
        return False

    log.info("Clicking Main (version=%d, prod_priority=1).", best_v)
    main_link = best_row.find_element(By.XPATH, ".//a[normalize-space()='Main']")
    js_click(driver, main_link)
    wait_spinner_gone(driver)
    log.info("Work Definition editor loading.")
    return True

def switch_to_main_frame(driver):
    frames = driver.find_elements(By.CSS_SELECTOR, "iframe")
    for f in frames:
        try:
            driver.switch_to.frame(f)
            if driver.find_elements(By.XPATH, "//*[contains(.,'kitting')]"):
                log.info("Switched into correct iframe.")
                return
            driver.switch_to.default_content()
        except:
            driver.switch_to.default_content()
    log.warning("No suitable iframe found.")

# ============================================================
#  STEP 7 – Double-click the "10 Kitting" operation
# ============================================================

def double_click_kitting(driver):
    log.info("Double-clicking Kitting…")

    label = find(driver, By.XPATH,
    "//*[name()='text' and contains(.,'10 Kitting')]"
    )

    kitting = label.find_element(By.XPATH, "./ancestor::*[name()='g'][1]")

    
    # =========================
    # ✅ METHOD 2: HIGHLIGHT
    # =========================
    driver.execute_script("""
        arguments[0].style.border='3px solid red';
        arguments[0].style.backgroundColor='rgba(255,0,0,0.2)';
    """, kitting)

    # =========================
    # ✅ METHOD 1: GET LOCATION
    # =========================
    rect = driver.execute_script("""
        var r = arguments[0].getBoundingClientRect();
        return {x: r.left, y: r.top, w: r.width, h: r.height};
    """, kitting)

    log.info(f"Kitting location: x={rect['x']}, y={rect['y']}, "
             f"width={rect['w']}, height={rect['h']}")

    time.sleep(2)  # 👈 pause so you can SEE the highlight

    # =========================
    # Existing double-click logic
    # =========================
    actions = ActionChains(driver)
    actions.move_to_element(kitting).pause(0.5).double_click().perform()

    # Step 3: scroll into view
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", kitting)
    time.sleep(1)

    # Step 4: get size and click CENTER via offset
    size = kitting.size
    width = size['width']
    height = size['height']

    actions = ActionChains(driver)

    try:
        actions.move_to_element_with_offset(kitting, width/2, height/2)\
               .pause(0.5)\
               .double_click()\
               .perform()
        log.info("Kitting double-click attempted (offset).")

    except Exception as e:
        log.warning(f"Offset double-click failed: {e}")
        log.info("Falling back to JS double-click.")

        driver.execute_script("""
            var rect = arguments[0].getBoundingClientRect();
            var x = rect.left + rect.width/2;
            var y = rect.top + rect.height/2;

            var evt = new MouseEvent('dblclick', {
                bubbles: true,
                cancelable: true,
                clientX: x,
                clientY: y
            });
            document.elementFromPoint(x, y).dispatchEvent(evt);
        """, kitting)

    wait_spinner_gone(driver)

# ============================================================
#  STEP 8+9+10 – Delete existing Items via Context Menu
# ============================================================

def delete_items_if_present(driver):
    log.info("Checking for Items sub-box…")

    items_box = maybe(driver, By.XPATH,
        "//*[normalize-space(text())='Items']"
        "[ancestor::*[contains(translate(.,'K','k'),'kitting')]]",
        timeout=5
    )

    if items_box is None:
        log.info("No Items box – skipping deletion.")
        return

    log.info("Items box found – initiating context menu deletion.")

    # 1. Click the 'Tick' icon on the top toolbar
    log.info("Clicking the top toolbar 'Tick' icon...")
    tick_icon = find(driver, By.XPATH, 
        "//a[@title='Select'] | //button[@title='Select'] | "
        "//img[contains(@src, 'select') or contains(@src, 'check')]/.. | "
        "//*[contains(@class,'Toolbar')]//*[contains(@class,'select') or contains(@class,'check')]"
    )
    js_click(driver, tick_icon)
    time.sleep(1)

    # 2. Click the circle box on the top right of "Items"
    log.info("Clicking the selection circle on the 'Items' box...")
    items_circle = find(driver, By.XPATH,
        "//*[normalize-space(text())='Items']/ancestor::*[name()='g'][1]//*[name()='circle']"
    )
    js_click(driver, items_circle)
    time.sleep(1)

    # 3. Right-click the basket inside "Items"
    log.info("Right-clicking the Items basket icon...")
    items_basket = find(driver, By.XPATH,
        "//*[normalize-space(text())='Items']/ancestor::*[name()='g'][1]//*[name()='image' or contains(@class,'icon')]"
    )
    ActionChains(driver).context_click(items_basket).perform()
    time.sleep(1)

    # 4. Click Actions -> Delete
    actions_menu = find(driver, By.XPATH, "//*[normalize-space(text())='Actions']")
    js_click(driver, actions_menu)
    time.sleep(0.5)

    delete_menu = find(driver, By.XPATH, "//*[normalize-space(text())='Delete' and not(ancestor::button)]")
    js_click(driver, delete_menu)
    time.sleep(1)

    # 5. Confirm delete dialog
    ok = clickable(driver, By.XPATH,
        "//div[@role='dialog']//button[normalize-space()='OK']"
        " | //button[normalize-space()='OK']",
        timeout=SHORT_WAIT
    )
    js_click(driver, ok)
    log.info("Delete dialog confirmed.")

    # Wait for Items box to disappear
    wait_gone(driver, By.XPATH,
        "//*[normalize-space(text())='Items']"
        "[ancestor::*[contains(translate(.,'K','k'),'kitting')]]",
        timeout=LONG_WAIT
    )
    wait_spinner_gone(driver)
    log.info("Items deleted.")

# ============================================================
#  STEP 12+13 – Right-click Item Structure node → Collect All Direct Children
# ============================================================

def collect_all_direct_children(driver):
    """
    Slide 7:
      Right-click the top rectangular item card in the right-hand
      'Item Structure: Primary' panel (shows the item name e.g. 'PANEL POSITION C').
      Select 'Collect All Direct Children' from the context menu.
    """
    log.info("Right-clicking Item Structure top card…")

    # The top card in the Item Structure panel on the right side of the page.
    # From slide 7: it is a rectangular box showing the item name with a coloured icon.
    # It sits directly below the panel header icons row.
    node = find(driver, By.XPATH,
        # Any element inside the Item Structure panel that looks like the top node
        "//*[contains(@id,'itemStructure') or contains(@id,'ItemStructure') or contains(@id,'bomStructure')]"
        "//*[contains(@class,'node') or contains(@class,'row') or contains(@class,'item')][1]"
        # Fallback: first dvt node (Oracle JET diagram node)
        " | //*[contains(@class,'oj-dvt') or contains(@class,'dvt-node')][1]"
    )

    ActionChains(driver).context_click(node).perform()
    wait_spinner_gone(driver, timeout=10)

    collect = clickable(driver, By.XPATH,
        "//*[normalize-space()='Collect All Direct Children']",
        timeout=SHORT_WAIT
    )
    js_click(driver, collect)
    wait_spinner_gone(driver)
    log.info("Collect All Direct Children done.")

# ============================================================
#  STEP 14+15 – Drag basket onto Kitting, wait, Save and Close
# ============================================================

def assign_items_to_kitting(driver):
    log.info("Assigning collected items via context menu…")

    # 1. Right-click the basket below item structure
    basket = find(driver, By.XPATH,
        "//*[contains(@id,'itemStructure') or contains(@id,'ItemStructure') or contains(@id,'bomStructure')]"
        "//*[contains(@class,'node') or contains(@class,'row') or contains(@class,'item')][1]"
        " | //*[contains(@class,'oj-dvt') or contains(@class,'dvt-node')][1]"
    )
    ActionChains(driver).context_click(basket).perform()
    time.sleep(1)

    # 2. Click Actions -> Assign
    actions_menu = find(driver, By.XPATH, "//*[normalize-space(text())='Actions']")
    js_click(driver, actions_menu)
    time.sleep(0.5)

    assign_menu = find(driver, By.XPATH, "//*[normalize-space(text())='Assign']")
    js_click(driver, assign_menu)
    time.sleep(1)

    # 3. Click OK in the popup window
    log.info("Confirming assignment popup...")
    ok = clickable(driver, By.XPATH,
        "//div[@role='dialog']//button[normalize-space()='OK']"
        " | //button[normalize-space()='OK']",
        timeout=SHORT_WAIT
    )
    js_click(driver, ok)
    wait_spinner_gone(driver, timeout=LONG_WAIT)

    # Confirm Items sub-box reappeared in Kitting
    try:
        W(driver, LONG_WAIT).until(EC.presence_of_element_located((
            By.XPATH,
            "//*[normalize-space(text())='Items']"
            "[ancestor::*[contains(translate(.,'K','k'),'kitting')]]"
        )))
        log.info("Assignment confirmed – Items visible in Kitting.")
    except TimeoutException:
        log.warning("Items box did not reappear – proceeding anyway.")

def save_and_close(driver):
    """Slide 8: Click Save and Close, wait for search page to return."""
    log.info("Save and Close…")
    btn = clickable(driver, By.XPATH,
        "//button[normalize-space()='Save and Close']"
        " | //input[@value='Save and Close']",
        timeout=SHORT_WAIT
    )
    js_click(driver, btn)

    log.info("Waiting for editor to close…")
    wait_gone(driver, By.XPATH,
        "//h1[contains(.,'Edit Work Definition')]"
        " | //span[contains(.,'Edit Work Definition Details')]",
        timeout=LONG_WAIT
    )
    wait_spinner_gone(driver, timeout=LONG_WAIT)

    find(driver, By.XPATH,
        "//h1[normalize-space()='Manage Work Definitions']"
        " | //span[normalize-space()='Manage Work Definitions']",
        timeout=MEDIUM_WAIT
    )
    log.info("Back on search page.")


def reset_search(driver):
    btn = maybe(driver, By.XPATH,
        "//button[normalize-space()='Reset'] | //input[@value='Reset']", timeout=5)
    if btn:
        js_click(driver, btn)
        wait_spinner_gone(driver, timeout=10)

# ============================================================
#  PER-ITEM ORCHESTRATION
# ============================================================

def run_with_retry(driver, item_number, summary):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            result = process_item(driver, item_number)
            (summary.success if result == "success" else summary.skipped).append(item_number)
            return
        except Exception as exc:
            log.warning("Attempt %d/%d failed for %s: %s",
                        attempt, MAX_RETRIES, item_number, exc,
                        exc_info=(attempt == MAX_RETRIES))
            if attempt < MAX_RETRIES:
                log.info("Retrying in 3s…")
                time.sleep(3)
                try:
                    cancel = maybe(driver, By.XPATH,
                                   "//button[normalize-space()='Cancel']", timeout=3)
                    if cancel:
                        js_click(driver, cancel)
                        wait_spinner_gone(driver, timeout=15)
                    if "Manage Work Definitions" not in driver.title:
                        open_manage_work_definitions(driver)
                except Exception:
                    pass
    log.error("Giving up on %s after %d attempts.", item_number, MAX_RETRIES)
    summary.failed.append(item_number)

def process_item(driver, item_number) -> str:
    log.info("─" * 50)
    log.info("Processing: %s", item_number)
    search_for_item(driver, item_number)
    if not click_main_link(driver, item_number):
        return "skipped"
    
    W(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(.,'Kitting')]"))
    )
    time.sleep(2) 

    switch_to_main_frame(driver)
    double_click_kitting(driver)

    W(driver, 30).until(
        EC.presence_of_element_located((
            By.XPATH,
            "//*[name()='text' and contains(.,'Items')]"
        ))
    )
    time.sleep(1)

    # Use the new context menu flow
    delete_items_if_present(driver)
    collect_all_direct_children(driver)
    assign_items_to_kitting(driver)  # Replaces drag_basket_to_kitting
    save_and_close(driver)
    reset_search(driver)
    
    log.info("✓  %s complete.", item_number)
    return "success"

# ============================================================
#  MAIN
# ============================================================

def main():
    try:
        items = load_items()
    except FileNotFoundError as exc:
        log.error("%s", exc)
        input("\nPress Enter to close…")
        sys.exit(1)

    if not items:
        log.error("No items found. Check EXCEL_FILENAME / EXCEL_COLUMN.")
        input("\nPress Enter to close…")
        sys.exit(1)

    summary = RunSummary()
    driver: Optional[webdriver.Firefox] = None
    try:
        driver = build_driver()
        navigate_and_login(driver)
        open_manage_work_definitions(driver)
        for idx, item in enumerate(items, 1):
            log.info("Item %d / %d", idx, len(items))
            run_with_retry(driver, item, summary)
    except KeyboardInterrupt:
        log.warning("Interrupted.")
    except Exception as exc:
        log.critical("Fatal: %s", exc, exc_info=True)
    finally:
        summary.report()
        if driver:
            driver.quit()


if __name__ == "__main__":
    main()

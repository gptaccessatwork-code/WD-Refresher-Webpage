"""
Oracle Cloud – Work Definition Refresh Automation
==================================================
Place this script and items.xlsx in the same folder, then run from VSCode.

Credentials are read from environment variables — never hardcoded:
    ORACLE_USER   your Oracle Cloud username / email
    ORACLE_PASS   your Oracle Cloud password

Set these in Windows:
    System Properties → Environment Variables → User Variables → New

Procedure per item
------------------
1.  Navigate → auto-login → Manage Work Definitions → search → click Main
2.  Double-click the first operation (Kitting, any sequence number)
3.  If Items sub-box is present:
      a. Click 'Done Collect' tick icon  (left toolbar)
      b. Click 'Not Collected' circle    (Items sub-box, located by SVG coordinates)
      c. Right-click left basket → Actions → Delete → OK
      d. Wait for Items to vanish
4.  Right-click top card in Item Structure panel → Collect All Direct Children
5.  Right-click right-panel basket → Actions → Assign → OK → wait for dialog
6.  Verify: click Expand All → count bold 'Items' nodes; if ≠ 1 → flag for review
7.  Save and Close → reset search → repeat

Each step retries up to MAX_RETRIES times on failure.
Persistent step failure → RuntimeError → program terminates.
Every click is visually highlighted in the browser.
Exceptions exported to wd_exceptions.xlsx at the end of the run.
"""

import logging
import os
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

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
#  USER CONFIG  ← edit these if needed
# ============================================================
if getattr(sys, 'frozen', False):
    # Running as an .exe - look in the folder where the .exe is located
    _HERE = os.path.dirname(sys.executable)
else:
    # Running as a .py script - look in the script folder
    _HERE = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILENAME = os.path.join(_HERE, "items.xlsx")
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
    success:      list = field(default_factory=list)
    skipped:      list = field(default_factory=list)   # no Main / no ProdPriority=1
    needs_review: list = field(default_factory=list)   # list of {"Item":…, "Reason":…} dicts

    def report(self) -> None:
        log.info("=" * 60)
        log.info("RUN SUMMARY")
        log.info("  Successful   : %d", len(self.success))
        log.info("  Skipped      : %d", len(self.skipped))
        log.info("  Needs review : %d", len(self.needs_review))
        if self.skipped:      log.info("  Skipped items      : %s", self.skipped)
        if self.needs_review:
            log.info("  Needs-review items : %s",
                     [r["Item"] for r in self.needs_review])
        log.info("=" * 60)

    def export_exceptions(self) -> None:
        """Write skipped + needs_review items to wd_exceptions.xlsx."""
        rows = (
            [{"Item": i, "Reason": "Skipped – no qualifying row found"}
             for i in self.skipped] +
            list(self.needs_review)   # already {"Item": …, "Reason": …} dicts
        )
        if not rows:
            log.info("No exceptions to export.")
            return
        out = Path(__file__).parent / "wd_exceptions.xlsx"
        pd.DataFrame(rows).to_excel(out, index=False)
        log.info("Exceptions exported → %s  (%d row(s))", out, len(rows))

# ============================================================
#  DRIVER
# ============================================================
def build_driver() -> webdriver.Firefox:
    opts = FirefoxOptions()
    opts.binary_location = FIREFOX_BINARY
    if HEADLESS:
        opts.add_argument("--headless")
    opts.set_preference("devtools.console.stdout.content", False)
    svc = FirefoxService(
        executable_path=GeckoDriverManager().install(),
        log_path="geckodriver.log",
    )
    driver = webdriver.Firefox(service=svc, options=opts)
    driver.maximize_window()
    return driver

# ============================================================
#  WAIT HELPERS
# ============================================================
def _w(driver, timeout) -> WebDriverWait:
    return WebDriverWait(driver, timeout, poll_frequency=POLL)

def find(driver, by, loc, timeout=MEDIUM_WAIT):
    """Wait until element is visible, return it."""
    return _w(driver, timeout).until(EC.visibility_of_element_located((by, loc)))

def clickable(driver, by, loc, timeout=MEDIUM_WAIT):
    """Wait until element is clickable, return it."""
    return _w(driver, timeout).until(EC.element_to_be_clickable((by, loc)))

def maybe(driver, by, loc, timeout=5):
    """Return element if found within timeout, else None."""
    try:
        return _w(driver, timeout).until(EC.presence_of_element_located((by, loc)))
    except TimeoutException:
        return None

def wait_gone(driver, by, loc, timeout=LONG_WAIT) -> None:
    """Block until element is no longer present/visible."""
    _w(driver, timeout).until(EC.invisibility_of_element_located((by, loc)))

def wait_spinner_gone(driver, timeout=LONG_WAIT) -> None:
    """Block until all Oracle ADF busy-spinners have cleared."""
    selectors = [
        (By.CSS_SELECTOR, "div.AFBusyIndicator"),
        (By.CSS_SELECTOR, "div[class*='AFBlockingGlassPane']"),
        (By.CSS_SELECTOR, "div[class*='af_loadingIndicator']"),
        (By.XPATH,        "//*[contains(@class,'AFBusyIndicator')]"),
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

def retry_step(step_name: str, fn: Callable, *args, **kwargs):
    """
    Run fn(*args, **kwargs) up to MAX_RETRIES times with a 2 s gap.
    On total failure raises RuntimeError, which terminates the program.
    """
    last_exc: Optional[Exception] = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:
            last_exc = exc
            log.warning("Step '%s' attempt %d/%d failed: %s",
                        step_name, attempt, MAX_RETRIES, exc)
            if attempt < MAX_RETRIES:
                time.sleep(2)
    raise RuntimeError(
        f"Step '{step_name}' failed after {MAX_RETRIES} attempts. "
        f"Last error: {last_exc}"
    ) from last_exc

# ============================================================
#  VISUAL HIGHLIGHT
# ============================================================
# Single JS snippet stored once and reused — avoids repeated string allocation.
_HIGHLIGHT_JS = """
var el  = arguments[0];
var tag = el.tagName.toLowerCase();
var svgTags = new Set(['image','circle','rect','ellipse','path',
                       'text','line','polyline','polygon','g','use']);
if (svgTags.has(tag)) {
    var svg = el;
    while (svg && svg.tagName.toLowerCase() !== 'svg') { svg = svg.parentElement; }
    if (!svg) return;
    try {
        var bbox = el.getBBox();
        var ctm  = el.getCTM();
        var pt   = svg.createSVGPoint();
        pt.x = bbox.x; pt.y = bbox.y;
        var tl = pt.matrixTransform(ctm);
        pt.x = bbox.x + bbox.width; pt.y = bbox.y + bbox.height;
        var br = pt.matrixTransform(ctm);
        var r  = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        r.setAttribute('x',             tl.x);
        r.setAttribute('y',             tl.y);
        r.setAttribute('width',         Math.max(br.x - tl.x, 4));
        r.setAttribute('height',        Math.max(br.y - tl.y, 4));
        r.setAttribute('fill',          'rgba(255,0,0,0.40)');
        r.setAttribute('stroke',        'red');
        r.setAttribute('stroke-width',  '2');
        r.setAttribute('pointer-events','none');
        svg.appendChild(r);
        setTimeout(function(){ try{ svg.removeChild(r); }catch(e){} }, 900);
    } catch(e) {}
} else {
    var o = el.style.outline, b = el.style.backgroundColor;
    el.style.outline = '3px solid red'; el.style.backgroundColor = 'rgba(255,0,0,0.25)';
    setTimeout(function(){ el.style.outline = o; el.style.backgroundColor = b; }, 900);
}
"""

def highlight(driver, el, label: str = "") -> None:
    """
    Flash a red overlay on el for ~0.9 s and log its attributes.
    SVG elements get a <rect> overlay (CSS outline has no effect on SVG).
    HTML elements get a CSS outline + background tint.
    Never raises — highlight failures must not interrupt the main flow.
    """
    try:
        log.info("  ► [%-30s]  tag=%-8s  id=%s  title=%s  alt=%s  aria=%s",
                 label,
                 el.tag_name,
                 el.get_attribute("id")          or "–",
                 el.get_attribute("title")        or "–",
                 el.get_attribute("alt")          or "–",
                 el.get_attribute("aria-label")   or "–")
        driver.execute_script(_HIGHLIGHT_JS, el)
        time.sleep(0.6)
    except Exception:
        pass

# ============================================================
#  CLICK HELPERS
# ============================================================
def js_click(driver, el, label: str = "") -> None:
    """Highlight → scroll into view → JS click.
    Use for HTML elements. JS click bypasses ADF's onclick='return false'."""
    highlight(driver, el, label)
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center',inline:'center'});", el)
    driver.execute_script("arguments[0].click();", el)

def svg_click(driver, el, label: str = "") -> None:
    """Highlight → scroll into view → ActionChains click.
    Use for SVG elements — they have no .click() DOM method in Firefox."""
    highlight(driver, el, label)
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center',inline:'center'});", el)
    ActionChains(driver).move_to_element(el).click().perform()

def right_click(driver, el, label: str = "") -> None:
    """Highlight → scroll into view → right-click.
    Waits 1.5 s for Oracle's context menu to render."""
    highlight(driver, el, label)
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center',inline:'center'});", el)
    ActionChains(driver).context_click(el).perform()
    time.sleep(1.5)

# ============================================================
#  CONTEXT-MENU HELPERS
# ============================================================
def _click_menu_item(driver, text: str) -> None:
    """Click a <td> context-menu item by exact text."""
    item = clickable(driver, By.XPATH,
        f"//td[normalize-space()='{text}']", timeout=SHORT_WAIT)
    highlight(driver, item, f"menu:{text}")
    ActionChains(driver).move_to_element(item).pause(0.2).click().perform()
    time.sleep(0.5)

def _actions_then(driver, action_name: str) -> None:
    """
    Right-click context menu: Actions → [action_name] (e.g., Delete).

    Uses multiple selector strategies to handle different menu rendering states.
    """
    log.info("Finding Actions menu and clicking '%s'…", action_name)

    # Wait for the context menu to appear
    try:
        menu = _w(driver, MEDIUM_WAIT).until(
            EC.presence_of_element_located((By.XPATH, "//td[contains(text(),'Actions')]")))
        log.info("  ✓ Actions menu found")
    except TimeoutException:
        log.error("  ✗ Actions menu NOT found")
        log.info("  Dumping all visible TD elements with text:")
        try:
            all_tds = driver.find_elements(By.XPATH, "//td")
            for idx, td in enumerate(all_tds[:20]):
                log.info("    [%d] %s", idx, td.text[:50])
        except:
            pass
        raise

    # Click the Actions menu item
    js_click(driver, menu, "menu:Actions")
    wait_spinner_gone(driver, timeout=SHORT_WAIT)
    time.sleep(1.0)

    # Find the specific action (Delete, etc.) using multiple selectors
    log.info("  Finding '%s' action in menu…", action_name)

    selectors = [
        f"//td[contains(text(),'{action_name}')]",
        f"//td[normalize-space()='{action_name}']",
        f"//div[contains(text(),'{action_name}')]",
        f"//*[contains(text(),'{action_name}')]",
    ]

    action_element = None
    for selector_idx, selector in enumerate(selectors):
        try:
            log.info("    Trying selector %d: %s", selector_idx + 1, selector)
            action_element = _w(driver, SHORT_WAIT).until(
                EC.presence_of_element_located((By.XPATH, selector)))
            log.info("    ✓ Found with selector %d", selector_idx + 1)
            break
        except TimeoutException:
            log.info("    ✗ Selector %d failed", selector_idx + 1)
            continue

    if action_element is None:
        log.error("  ✗ '%s' action NOT found with any selector", action_name)
        log.info("  Dumping all visible menu items:")
        try:
            all_items = driver.find_elements(By.XPATH, "//td | //div[@role='menuitem']")
            for idx, item in enumerate(all_items[:30]):
                log.info("    [%d] Tag=%s, Text='%s'", idx, item.tag_name, item.text[:50])
        except:
            pass
        raise RuntimeError(f"Could not locate '{action_name}' in menu")

    # Click the action
    log.info("  Clicking '%s' action…", action_name)
    try:
        js_click(driver, action_element, f"menu:{action_name}")
        log.info("  ✓ '%s' clicked successfully", action_name)
    except Exception as e:
        log.error("  ✗ Failed to click '%s': %s", action_name, e)
        raise

    wait_spinner_gone(driver, timeout=SHORT_WAIT)

def _actions_then_original(driver, child: str) -> None:
    """
    Original v1.1 version for Assign.
    In the open context menu: click 'Actions' to reveal the sub-menu,
    wait for it to appear, then click the child item ('Assign').
    """
    _w(driver, SHORT_WAIT).until(
        EC.presence_of_element_located(
            (By.XPATH, "//td[normalize-space()='Actions']")))
    time.sleep(0.5)

    actions_td = clickable(driver, By.XPATH,
        "(//td[normalize-space()='Actions'"
        "      and not(contains(@style,'display: none'))])[last()]",
        timeout=SHORT_WAIT,
    )
    highlight(driver, actions_td, "menu:Actions")
    ActionChains(driver).move_to_element(actions_td).pause(0.2).click().perform()
    time.sleep(1.0)

    _click_menu_item(driver, child)

def _confirm_ok(driver, label: str = "OK") -> None:
    """Wait for an OK button and JS-click it."""
    ok = clickable(driver, By.XPATH,
        "//button[normalize-space()='OK']", timeout=SHORT_WAIT)
    js_click(driver, ok, label)
    wait_spinner_gone(driver, timeout=15)

# ============================================================
#  EXCEL LOADER
# ============================================================
def load_items() -> list:
    script_dir = Path(__file__).parent.resolve()
    path = script_dir / EXCEL_FILENAME
    if not path.exists():
        cands = (list(script_dir.glob("*.xlsx")) +
                 list(script_dir.glob("*.xls")) +
                 list(script_dir.glob("*.csv")))
        raise FileNotFoundError(
            f"\nCannot find '{EXCEL_FILENAME}' in:\n  {script_dir}\n"
            f"Files found: {[f.name for f in cands]}\n"
            f"Rename your file to '{EXCEL_FILENAME}' or update EXCEL_FILENAME."
        )
    suffix = path.suffix.lower()
    df = (pd.read_excel(path, sheet_name=EXCEL_SHEET, header=None, dtype=str)
          if suffix in (".xlsx", ".xls")
          else pd.read_csv(path, header=None, dtype=str))
    col = (ord(EXCEL_COLUMN.upper()) - ord("A")
           if isinstance(EXCEL_COLUMN, str) else int(EXCEL_COLUMN))
    items = (df.iloc[:, col].dropna().str.strip()
               .replace("", pd.NA).dropna().unique().tolist())
    log.info("Loaded %d item(s) from '%s'.", len(items), path.name)
    return items

# ============================================================
#  STEP 1 – Navigate and auto-login
# ============================================================
def _get_credentials() -> tuple[str, str]:
    """
    Read Oracle credentials from environment variables.
    Exits with a clear message if either is missing.
    """
    user = os.environ.get("ORACLE_USER", "").strip()
    pwd  = os.environ.get("ORACLE_PASS", "").strip()
    if not user or not pwd:
        missing = [v for v, val in [("ORACLE_USER", user), ("ORACLE_PASS", pwd)] if not val]
        log.error(
            "Missing environment variable(s): %s\n"
            "Set them in Windows → System Properties → Environment Variables → "
            "User Variables → New, then restart VSCode.",
            ", ".join(missing),
        )
        sys.exit(1)
    return user, pwd

def navigate_and_login(driver) -> None:
    """
    Navigate to Oracle Cloud and log in automatically using ORACLE_USER /
    ORACLE_PASS environment variables.

    Oracle Fusion's login is a two-step form:
      Step 1 – enter username, click Next / Continue
      Step 2 – enter password, click Sign In

    If the login page is not detected (e.g. session already active),
    the function returns immediately.
    """
    log.info("Opening Oracle Cloud…")
    driver.get(ORACLE_URL)
    wait_spinner_gone(driver, timeout=30)

    # If no password field is visible, we're already logged in
    if maybe(driver, By.CSS_SELECTOR, "input[type='password']", timeout=5) is None:
        # Check for username field (step 1 of login)
        username_field = maybe(driver, By.CSS_SELECTOR,
            "input[type='text'], input[type='email']", timeout=5)
        if username_field is None:
            log.info("Already logged in.")
            return

    user, pwd = _get_credentials()
    log.info("Login page detected – logging in automatically…")

    # ── Step 1: username ────────────────────────────────────────
    username_field = maybe(driver, By.CSS_SELECTOR,
        "input[type='text'], input[type='email']", timeout=10)
    if username_field:
        highlight(driver, username_field, "Username field")
        username_field.clear()
        username_field.send_keys(user)
        # Click the Next / Continue button to advance to the password step
        next_btn = maybe(driver, By.XPATH,
            "//input[@type='submit']"
            " | //button[contains(translate(normalize-space(),"
            "  'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'next')]"
            " | //button[contains(translate(normalize-space(),"
            "  'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'continue')]",
            timeout=5,
        )
        if next_btn:
            js_click(driver, next_btn, "Next / Continue")
            wait_spinner_gone(driver, timeout=15)

    # ── Step 2: password ────────────────────────────────────────
    password_field = find(driver, By.CSS_SELECTOR,
        "input[type='password']", timeout=MEDIUM_WAIT)
    highlight(driver, password_field, "Password field")
    password_field.clear()
    password_field.send_keys(pwd)

    sign_in_btn = clickable(driver, By.XPATH,
        "//input[@type='submit']"
        " | //button[contains(translate(normalize-space(),"
        "  'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'sign in')]"
        " | //button[contains(translate(normalize-space(),"
        "  'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login')]",
        timeout=SHORT_WAIT,
    )
    js_click(driver, sign_in_btn, "Sign In")

    # Wait until the login form disappears
    wait_gone(driver, By.CSS_SELECTOR, "input[type='password']", timeout=60)
    wait_spinner_gone(driver, timeout=30)
    log.info("Login complete.")

# ============================================================
#  STEP 2 – Open Manage Work Definitions
# ============================================================
def open_manage_work_definitions(driver) -> None:
    log.info("Opening Manage Work Definitions…")

    star = clickable(driver, By.XPATH,
        "//*[contains(@id,'pt1') and contains(@id,'star')]"
        " | //*[@title='Favorites and Recent Items']"
        " | //*[contains(@class,'oj-ux-ico-star')]/.."
        " | //img[contains(@src,'star') or contains(@src,'favorite')]/..",
    )
    js_click(driver, star, "Favorites star")
    wait_spinner_gone(driver, timeout=10)

    mwd = clickable(driver, By.XPATH,
        "//a[normalize-space()='Manage Work Definitions']"
        " | //span[normalize-space()='Manage Work Definitions']",
    )
    js_click(driver, mwd, "Manage Work Definitions")
    wait_spinner_gone(driver)

    find(driver, By.XPATH,
        "//h1[normalize-space()='Manage Work Definitions']"
        " | //*[contains(@id,'ManageWorkDef')]",
    )
    log.info("Manage Work Definitions loaded.")

# ============================================================
#  STEP 3+4 – Search for item
# ============================================================
def _do_search(driver, item_number: str) -> None:
    wait_spinner_gone(driver, timeout=15)

    # Wait for the Item input to be enabled (prevents SPA race condition)
    _w(driver, MEDIUM_WAIT).until(
        lambda d: d.find_element(
            By.XPATH, "//input[contains(@id,'qryId1:value00')]"
        ).is_enabled()
    )
    time.sleep(0.5)

    # Locate and interact with the Item field
    item_input = find(driver, By.XPATH,
        "//input[contains(@id,'qryId1:value00')]")
    highlight(driver, item_input, "Item field")
    item_input.click()
    time.sleep(0.2)
    item_input.send_keys(Keys.CONTROL + "a")
    item_input.send_keys(Keys.DELETE)
    item_input.send_keys(item_number)
    time.sleep(0.3)

    # Search button id always ends with '::search'
    search_btn = clickable(driver, By.XPATH,
        "//button[substring(@id, string-length(@id) - 7) = '::search']"
        " | //button[normalize-space()='Search']",
    )
    js_click(driver, search_btn, "Search button")
    wait_spinner_gone(driver)

def search_for_item(driver, item_number: str) -> None:
    log.info("Searching for: %s", item_number)
    retry_step("search", _do_search, driver, item_number)

# ============================================================
#  STEP 5+6 – Select best result row and click Main
# ============================================================
def click_main_link(driver, item_number: str) -> bool:
    """
    Pick the row for item_number where Production Priority = 1 and Version
    is highest, then click its 'Main' link.
    Returns False to skip this item if no qualifying row is found.
    """
    log.info("Waiting for search results…")
    row_xpath = (f"//tr[contains(.,'{item_number}')"
                 f"    and .//a[normalize-space()='Main']]")
    try:
        _w(driver, MEDIUM_WAIT).until(
            EC.presence_of_element_located((By.XPATH, row_xpath)))
    except TimeoutException:
        log.warning("No results for %s – skipping.", item_number)
        return False

    rows = driver.find_elements(By.XPATH, row_xpath)
    if not rows:
        log.warning("No rows with 'Main' link for %s – skipping.", item_number)
        return False

    def digits(s: str) -> int:
        """Extract the first integer from a string; -1 if none found."""
        d = "".join(c for c in s if c.isdigit())
        return int(d) if d else -1

    def span_x1a_int(cell) -> int:
        """
        Oracle renders editable numeric cells as:
            <td>
              <span class="x1a">1
                <img alt="">
                <a title="Edit Priorities">...</a>
              </span>
            </td>

        .text / textContent both pull in the child img alt and anchor text,
        corrupting the number. The only reliable read is the firstChild text
        node of the x1a span, accessed via JavaScript.
        """
        result = driver.execute_script("""
            var spans = arguments[0].querySelectorAll('span.x1a');
            if (!spans.length) return null;
            var node = spans[0].firstChild;
            return node ? node.textContent.trim() : null;
        """, cell)
        if result is not None:
            return digits(result)
        # Fallback: strip every non-digit character from the whole cell text
        return digits(cell.text.strip())

    best_row, best_v = None, -1
    for row in rows:
        try:
            cells  = row.find_elements(By.TAG_NAME, "td")
            texts  = [c.text.strip() for c in cells]
            offset = 1 if (texts and texts[0] == "") else 0

            # Try @headers-based column lookup first (most reliable when present).
            # Fall back to positional index when headers are absent.
            vc  = row.find_elements(By.XPATH,
                ".//td[contains(@headers,'version') or contains(@headers,'Version')]")
            ppc = row.find_elements(By.XPATH,
                ".//td[contains(@headers,'ProductionPriority')"
                "   or contains(@headers,'productionPriority')"
                "   or contains(@headers,'production')"
                "   or contains(@headers,'prodPriority')"
                "   or contains(@headers,'ProdPriority')]")

            v  = span_x1a_int(vc[0])  if vc  else span_x1a_int(cells[5 + offset]) if len(cells) > 5 + offset else -1
            pp = span_x1a_int(ppc[0]) if ppc else span_x1a_int(cells[6 + offset]) if len(cells) > 6 + offset else -1

            log.debug("Row cells: %s  → version=%d  prod_priority=%d", texts, v, pp)
            if pp == 1 and v > best_v:
                best_v, best_row = v, row
        except (IndexError, StaleElementReferenceException):
            continue

    if best_row is None:
        log.warning("No row with ProdPriority=1 for %s – skipping.", item_number)
        return False

    log.info("Opening Main (version=%d, prod_priority=1).", best_v)

    def _click():
        link = best_row.find_element(By.XPATH, ".//a[normalize-space()='Main']")
        js_click(driver, link, "Main link")
        wait_spinner_gone(driver)
        find(driver, By.XPATH,
            "//h1[contains(.,'Edit Work Definition')]"
            " | //span[contains(.,'Edit Work Definition Details')]",
            timeout=MEDIUM_WAIT,
        )

    retry_step("click_main", _click)
    log.info("Work Definition editor loaded.")
    return True

# ============================================================
#  STEP 7 – Double-click the first operation (Kitting)
# ============================================================
def _do_double_click_kitting(driver) -> None:
    find(driver, By.XPATH, "//*[name()='svg']", timeout=MEDIUM_WAIT)

    label = find(driver, By.XPATH,
        "//*[name()='text'][contains("
        "  translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "  'abcdefghijklmnopqrstuvwxyz'),'kitting')]",
        timeout=MEDIUM_WAIT,
    )
    node = label.find_element(By.XPATH, "./ancestor::*[name()='g'][1]")

    highlight(driver, node, "Kitting operation node")
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center',inline:'center'});", node)
    time.sleep(0.3)

    try:
        ActionChains(driver).double_click(node).perform()
    except Exception:
        driver.execute_script("""
            var el = arguments[0], r = el.getBoundingClientRect();
            var cx = r.left + r.width/2, cy = r.top + r.height/2;
            (document.elementFromPoint(cx,cy)||el).dispatchEvent(
                new MouseEvent('dblclick',
                    {bubbles:true,cancelable:true,clientX:cx,clientY:cy}));
        """, node)

    _w(driver, SHORT_WAIT).until(EC.presence_of_element_located((By.XPATH,
        "//*[name()='text']"
        "[normalize-space()='Items' or normalize-space()='Resources']"
    )))

def double_click_kitting(driver) -> None:
    log.info("Double-clicking Kitting operation…")
    retry_step("double_click_kitting", _do_double_click_kitting, driver)
    log.info("Kitting expanded.")

# ============================================================
#  EXPAND ALL OPERATIONS
# ============================================================
def _expand_all_operations(driver) -> None:
    """Click Expand All to show all operation cards."""
    log.info("Clicking Expand All button…")
    btn = find(driver, By.XPATH, _EXPAND_ALL_BTN, timeout=SHORT_WAIT)
    js_click(driver, btn, "Expand All")
    wait_spinner_gone(driver, timeout=SHORT_WAIT)
    time.sleep(3.0)
    log.info("All operations expanded.")

def expand_all_operations(driver) -> None:
    """Public wrapper."""
    retry_step("expand_all_operations", _expand_all_operations, driver)

# ============================================================
#  SELECTORS  (confirmed from live DOM inspection)
# ============================================================

# SVG <text> label for the Items sub-box inside Kitting
_ITEMS_SVG_TEXT = "//*[name()='text'][normalize-space()='Items']"

# 'Done Collect' tick icon in the left toolbar
_TICK_ICON = (
    "//img[@title='Done Collect' or @alt='Done Collect']"
    " | //img[contains(@src,'func_selectcheck_24_act')]"
    " | //*[contains(@id,'sModBt')]"
)

# Left-toolbar basket (delete flow) — id contains 'r1:0:basketImgLink'
_BASKET_LEFT = (
    "//img[contains(@id,'r1:0:basketImgLink')]"
    " | //img[@title='Collected Items' and contains(@id,'r1:')]"
    " | //img[contains(@src,'func_basket_24_ena') and contains(@id,'r1:')]"
)

# Right-panel basket (assign flow) — id contains 'r3:0:basketImgLink'
_BASKET_RIGHT = (
    "//img[contains(@id,'r3:0:basketImgLink')]"
    " | //img[@title='Collected Items' and contains(@id,'r3:')]"
    " | //img[contains(@src,'func_basket_24_ena') and contains(@id,'r3:')]"
)

# "Expand All" icon in the left canvas toolbar — used before verification.
# <img id="...:cil02::icon" title="Expand All" alt="Expand All"
#      src="…/func_expand_24_ena.png">
_EXPAND_ALL_BTN = (
    "//img[@title='Expand All' or @alt='Expand All']"
    " | //img[contains(@src,'func_expand_24_ena')]"
    " | //*[contains(@id,'cil02')]"
)

# After Expand All, count every SVG <text> node whose text is exactly "Items".
# This is the same simple selector used by _ITEMS_SVG_TEXT (which already works
# reliably throughout the script), applied here as a count after Expand All.
# Using only name()='text' + text content avoids fragility from attribute
# constraints — Oracle does not guarantee all styling attributes are present
# on every rendered operation node.
_ITEMS_COUNT_XPATH = "//*[name()='text'][normalize-space()='Items']"

_CANCEL_BTN     = "//button[normalize-space()='Cancel']"
_SAVE_CLOSE_BTN = "//button[normalize-space()='Save and Close']"

# ============================================================
#  SVG COORDINATE HELPER – find Items circle by proximity
# ============================================================
# JavaScript executed once per call; stored as a module-level constant
# to avoid rebuilding the string on every invocation.
_FIND_ITEMS_CIRCLE_JS = """
    var allText = document.querySelectorAll('text');
    var itemsEl = null;
    for (var i = 0; i < allText.length; i++) {
        if (allText[i].textContent.trim() === 'Items') {
            itemsEl = allText[i]; break;
        }
    }
    if (!itemsEl) return null;

    var itemsY  = itemsEl.getBoundingClientRect();
    itemsY = itemsY.top + itemsY.height / 2;

    var circles = document.querySelectorAll(
        'image[*|href*="qual_selectcheckblank_24"]');
    if (!circles.length) return null;

    var best = null, bestDelta = Infinity;
    for (var j = 0; j < circles.length; j++) {
        var r  = circles[j].getBoundingClientRect();
        var cy = r.top + r.height / 2;
        var d  = Math.abs(cy - itemsY);
        if (d < bestDelta) { bestDelta = d; best = circles[j]; }
    }
    return best;
"""

def _find_items_circle(driver):
    """
    Return the 'Not Collected' SVG circle belonging to the Items sub-box.

    XPath cannot reliably distinguish the Items circle from the Resources
    circle because their DOM nesting varies. Instead we use JS to find all
    qual_selectcheckblank_24 images and return the one whose viewport
    Y-centre is closest to the 'Items' text label — coordinate-based,
    nesting-independent.
    """
    el = driver.execute_script(_FIND_ITEMS_CIRCLE_JS)
    if el is None:
        raise RuntimeError(
            "Could not locate 'Not Collected' circle near Items label. "
            "Confirm Kitting is expanded and the Items sub-box is visible."
        )
    return el

# ============================================================
#  STEP 8-10 – Delete existing Items (if present)
# ============================================================
def _do_delete_items(driver) -> None:
    # Step 9: click 'Done Collect' tick icon.
    # After clicking, Oracle redraws the SVG canvas.  We wait for the tick's
    # own title to change from "Collect" → "Done Collect" as the definitive
    # signal that the canvas is ready — then add an extra buffer.
    log.info("Step 9: clicking 'Done Collect' tick…")
    tick = find(driver, By.XPATH, _TICK_ICON, timeout=SHORT_WAIT)
    js_click(driver, tick, "Done Collect tick")
    wait_spinner_gone(driver, timeout=SHORT_WAIT)

    # Wait until the tick's title reads "Done Collect" (canvas fully redrawn)
    log.info("Step 9: waiting for canvas to settle after tick click…")
    _w(driver, SHORT_WAIT).until(
        lambda d: (d.find_elements(By.XPATH,
            "//img[@title='Done Collect' or @alt='Done Collect']"
            " | //a[@title='Done Collect']"
            " | //*[contains(@id,'sModBt') and @title='Done Collect']")
        ) or True  # fallback: just wait the full sleep below
    )
    wait_spinner_gone(driver, timeout=SHORT_WAIT)
    time.sleep(3.0)   # extra buffer: SVG circles render asynchronously

    # Step 9a: click 'Not Collected' circle on the Items sub-box.
    # Located by JS coordinate proximity to the 'Items' SVG text label.
    log.info("Step 9a: clicking 'Not Collected' circle on Items…")
    circle = _find_items_circle(driver)
    svg_click(driver, circle, "Not Collected circle (Items)")
    wait_spinner_gone(driver, timeout=SHORT_WAIT)
    time.sleep(3.0)   # Oracle processes selection; basket must be ready before right-click

    # Log circle state for diagnostics
    circle = _find_items_circle(driver)
    state  = (circle.get_attribute("aria-label") or
              circle.get_attribute("href")        or
              circle.get_attribute("xlink:href"))
    log.info("Circle state after click: %s", state)

    # Step 9b: right-click left basket → Actions → Delete
    log.info("Step 9b: right-clicking left basket → Actions → Delete…")
    basket = find(driver, By.XPATH, _BASKET_LEFT, timeout=SHORT_WAIT)
    right_click(driver, basket, "Left basket")
    _actions_then(driver, "Delete")

    # Step 10: wait for the Delete confirmation dialog, click OK, wait for Items to vanish
    log.info("Step 10: waiting for Delete confirmation dialog…")
    _w(driver, MEDIUM_WAIT).until(
        EC.presence_of_element_located((By.XPATH, "//button[normalize-space()='OK']")))
    _confirm_ok(driver, "Delete OK")
    log.info("Waiting for Items sub-box to disappear…")
    wait_gone(driver, By.XPATH, _ITEMS_SVG_TEXT, timeout=LONG_WAIT)
    wait_spinner_gone(driver)
    log.info("Items deleted.")

def delete_items_if_present(driver) -> None:
    """
    Delete items sequentially, one operation at a time.

    Process:
    1. Find first Items (operations already expanded by caller)
    2. Tick → Circle → Delete it
    3. Loop back to find next Items
    4. Repeat until none remain
    """
    max_iterations = 10
    iteration = 0

    # Click Done Collect tick once - it stays clicked across operations
    log.info("Step 9: clicking 'Done Collect' tick…")
    tick = find(driver, By.XPATH, _TICK_ICON, timeout=SHORT_WAIT)
    js_click(driver, tick, "Done Collect tick")
    wait_spinner_gone(driver, timeout=SHORT_WAIT)
    time.sleep(3.0)

    # Wait for SVG canvas to be fully rendered after tick click
    find(driver, By.XPATH, "//*[name()='svg']", timeout=MEDIUM_WAIT)
    time.sleep(1.0)

    while iteration < max_iterations:
        iteration += 1
        log.info("Scan iteration %d…", iteration)

        # Find first remaining Items
        items = maybe(driver, By.XPATH, _ITEMS_SVG_TEXT, timeout=5)
        if items is None:
            log.info("No Items sub-box remaining – deletion complete.")
            return

        log.info("Items sub-box found – deleting.")

        # DELETE THIS OPERATION'S ITEMS
        log.info("Step 9a: clicking 'Not Collected' circle on Items…")
        circle = _find_items_circle(driver)
        svg_click(driver, circle, "Not Collected circle (Items)")
        wait_spinner_gone(driver, timeout=SHORT_WAIT)
        time.sleep(3.0)

        log.info("Step 9b: right-clicking left basket → Actions → Delete…")
        basket = find(driver, By.XPATH, _BASKET_LEFT, timeout=SHORT_WAIT)
        right_click(driver, basket, "Left basket")
        _actions_then(driver, "Delete")

        log.info("Step 10: Delete confirmation…")
        _w(driver, MEDIUM_WAIT).until(
            EC.presence_of_element_located((By.XPATH, "//button[normalize-space()='OK']")))
        _confirm_ok(driver, "Delete OK")

        # Wait for the Delete dialog to disappear
        log.info("Waiting for Delete dialog to close…")
        wait_gone(driver, By.XPATH, "//div[contains(@class,'x1jo') and normalize-space()='Delete']", timeout=LONG_WAIT)
        wait_spinner_gone(driver, timeout=LONG_WAIT)
        time.sleep(1.0)
        log.info("Items deleted from this operation.")

    log.warning("Reached max iterations (%d)", max_iterations)

# ============================================================
#  STEP 12-13 – Collect All Direct Children
# ============================================================
def _do_collect_all_direct_children(driver) -> None:
    top_card = find(driver, By.XPATH,
        "(//*[name()='rect'][@fill='#ffffff']"
        "    [contains(@stroke,'#5b74b7') or contains(@stroke,'5b74b7')])[1]",
        timeout=MEDIUM_WAIT,
    )
    right_click(driver, top_card, "Item Structure top card")

    collect = clickable(driver, By.XPATH,
        "//td[normalize-space()='Collect All Direct Children']",
        timeout=SHORT_WAIT,
    )
    js_click(driver, collect, "Collect All Direct Children")
    wait_spinner_gone(driver)

    _w(driver, MEDIUM_WAIT).until(
        EC.presence_of_element_located((By.XPATH, _BASKET_RIGHT)))
    log.info("Right-panel basket populated.")

def collect_all_direct_children(driver) -> None:
    log.info("Collecting all direct children…")
    retry_step("collect_all_direct_children",
               _do_collect_all_direct_children, driver)
    log.info("Collection complete.")

# ============================================================
#  STEP 14-15 – Assign basket to Kitting
# ============================================================

# The Assign Operation Items dialog title element, confirmed from live DOM:
#   <div id="...:pw2::_ttxt" class="x1jo">Assign Operation Items</div>
# This is the most reliable signal the dialog is open.
_ASSIGN_DIALOG_TITLE = (
    "//div[contains(@id,'pw2::_ttxt')"
    "      and normalize-space()='Assign Operation Items']"
    # Fallback: any element with class x1jo containing the title text
    " | //div[contains(@class,'x1jo')"
    "         and normalize-space()='Assign Operation Items']"
)

# The OK button inside the Assign dialog, confirmed from live DOM:
#   <button id="...:r22:0:cb3" ...>OK</button>
# Scoped inside the dialog container to avoid matching other OK buttons.
_ASSIGN_DIALOG_OK = (
    "//button[contains(@id,':r22:0:cb3')]"
    # Fallback: OK button that is a descendant of the dialog container
    " | //div[contains(@id,':r22')]//button[normalize-space()='OK']"
)

def _do_assign_basket(driver) -> None:
    log.info("Step 14: right-clicking right-panel basket…")
    basket = find(driver, By.XPATH, _BASKET_RIGHT, timeout=SHORT_WAIT)
    right_click(driver, basket, "Right basket")
    _actions_then_original(driver, "Assign")

    # Wait for the Assign Operation Items dialog title to appear in the DOM.
    # This is the definitive signal that the dialog is fully open.
    log.info("Step 15: waiting for 'Assign Operation Items' dialog to open…")
    _w(driver, MEDIUM_WAIT).until(
        EC.presence_of_element_located((By.XPATH, _ASSIGN_DIALOG_TITLE)))
    log.info("Step 15: dialog open — clicking OK…")

    # Click the dialog's own OK button (scoped by id pattern :r22:0:cb3)
    ok_btn = find(driver, By.XPATH, _ASSIGN_DIALOG_OK, timeout=SHORT_WAIT)
    js_click(driver, ok_btn, "Assign dialog OK")
    wait_spinner_gone(driver, timeout=15)

    # Wait for the dialog title to disappear — dialog closed = assignment done
    log.info("Waiting for Assign dialog to close…")
    wait_gone(driver, By.XPATH, _ASSIGN_DIALOG_TITLE, timeout=LONG_WAIT)
    wait_spinner_gone(driver, timeout=LONG_WAIT)
    log.info("Assignment complete.")

def assign_basket_to_kitting(driver) -> None:
    log.info("Assigning basket to Kitting…")
    retry_step("assign_basket", _do_assign_basket, driver)

# ============================================================
#  POST-ASSIGN VERIFICATION
# ============================================================
def verify_assignment(driver, item_number: str) -> Optional[str]:
    """
    Verification procedure:
      1. Click the 'Expand All' icon in the left canvas toolbar so that
         every operation's sub-panels (Items, Resources) are fully visible.
      2. Count occurrences of the bold 'Items' SVG text node.
         Exactly 1 → correct (one Kitting operation has its Items populated).
         More than 1 → assignment appears duplicated across operations; flag.
         Zero → Items panel not found after expand; flag.

    Returns None on success, or a reason string to write to wd_exceptions.xlsx.
    """
    log.info("Verification: clicking 'Expand All' to reveal all sub-panels…")
    expand_btn = maybe(driver, By.XPATH, _EXPAND_ALL_BTN, timeout=SHORT_WAIT)
    if expand_btn:
        js_click(driver, expand_btn, "Expand All")
        wait_spinner_gone(driver, timeout=SHORT_WAIT)
        time.sleep(1.0)   # let the SVG re-render after expansion
    else:
        log.warning("'Expand All' button not found – proceeding without it.")

    log.info("Verification: counting bold 'Items' text nodes in canvas…")
    items_nodes = driver.find_elements(By.XPATH, _ITEMS_COUNT_XPATH)
    count = len(items_nodes)
    log.info("Bold 'Items' node count: %d", count)

    if count == 1:
        log.info("✓ Verification passed for %s (Items count = 1).", item_number)
        return None

    if count == 0:
        reason = "Bold 'Items' text not found after Expand All – assignment state unknown"
    else:
        reason = (f"Bold 'Items' text found {count} time(s) after Expand All – "
                  f"WD may be over-assigned or assigned to multiple operations")

    log.warning("%s – %s", item_number, reason)
    return reason

# ============================================================
#  STEP 16 – Save and Close
# ============================================================
def _do_save_and_close(driver) -> None:
    btn = clickable(driver, By.XPATH, _SAVE_CLOSE_BTN, timeout=SHORT_WAIT)
    js_click(driver, btn, "Save and Close")
    wait_gone(driver, By.XPATH,
        "//h1[contains(.,'Edit Work Definition')]"
        " | //span[contains(.,'Edit Work Definition Details')]",
        timeout=LONG_WAIT,
    )
    wait_spinner_gone(driver, timeout=LONG_WAIT)
    find(driver, By.XPATH,
        "//h1[normalize-space()='Manage Work Definitions']"
        " | //span[normalize-space()='Manage Work Definitions']",
        timeout=MEDIUM_WAIT,
    )

def save_and_close(driver) -> None:
    log.info("Save and Close…")
    retry_step("save_and_close", _do_save_and_close, driver)
    log.info("Returned to search page.")
    wait_spinner_gone(driver, timeout=10)
    time.sleep(1.0)

# ============================================================
#  PER-ITEM ORCHESTRATION
# ============================================================
def process_item(driver, item_number: str) -> tuple:
    """
    Returns:
      ("success",      None)        – fully assigned
      ("skipped",      None)        – no qualifying WD row found
      ("needs_review", reason_str)  – saved but bad assignment detected
    """
    log.info("─" * 50)
    log.info("Processing: %s", item_number)

    search_for_item(driver, item_number)

    if not click_main_link(driver, item_number):
        return "skipped", None

    # Wait for the SVG canvas to be fully loaded before expanding
    find(driver, By.XPATH, "//*[name()='svg']", timeout=MEDIUM_WAIT)
    time.sleep(1.0)
    expand_all_operations(driver)
    delete_items_if_present(driver)
    collect_all_direct_children(driver)
    assign_basket_to_kitting(driver)

    bad_reason = verify_assignment(driver, item_number)

    # Always save and close — even when assignment is bad
    save_and_close(driver)
    open_manage_work_definitions(driver)

    if bad_reason is None:
        log.info("✓  %s complete.", item_number)
        return "success", None

    log.warning("⚠  %s saved but flagged for review.", item_number)
    return "needs_review", bad_reason

# ============================================================
#  MAIN
# ============================================================
def main() -> None:
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
            result, reason = process_item(driver, item)
            if result == "success":
                summary.success.append(item)
            elif result == "needs_review":
                summary.needs_review.append({"Item": item, "Reason": reason})
            else:
                summary.skipped.append(item)

    except RuntimeError as exc:
        log.critical("FATAL – step failure: %s", exc)
        log.critical("Terminating. Fix the issue and re-run from the failed item.")
    except KeyboardInterrupt:
        log.warning("Interrupted by user.")
    except Exception as exc:
        log.critical("Unexpected error: %s", exc, exc_info=True)
    finally:
        summary.report()
        summary.export_exceptions()
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()

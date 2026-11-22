"""
TWAMP BIH + JHK  -  PARALLEL DOWNLOAD + AUTO CONSOLIDATION  (v2)

Flow
  1. Two INDEPENDENT Microsoft Edge sessions start together (one process each).
  2. Each logs in, opens the 24 Hour Packet Loss Report, selects ONE circle, applies,
     waits for data, exports Excel.
  3. Main process waits until BOTH downloads finish.
  4. BIH + JHK are merged (streaming, low memory) into ONE workbook, validated,
     and installed as FINAL_SOURCE. BIH/JHK staging files are then deleted.

Requirements : pip install playwright openpyxl
Run          : python twamp_parallel_v2.py
credentials.txt beside this script:
    user=YOUR_USERNAME
    password=YOUR_PASSWORD
"""
from pathlib import Path
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed
import multiprocessing
import os, re, shutil, sys, time, traceback

# ----------------------------------------------------------------------------- settings
MYCOM_URL     = "https://welcome-page-pm-eaaprod.apps.mycom1.airtel.com/"
CIRCLES       = ("BIH", "JHK")
STAGING_DIR   = Path(r"C:\TWAMP_Automation\Staging")
BACKUP_DIR    = Path(r"C:\TWAMP_Automation\Backups")
FAIL_DIR      = Path(r"C:\TWAMP_Automation\Failures")
LOCK_FILE     = Path(r"C:\TWAMP_Automation\run.lock")
FINAL_SOURCE  = Path(r"E:\BOT\Power Query\Raw Dump\24 Hour Packet Loss Report_web.xlsx")
TEMP_SOURCE   = FINAL_SOURCE.with_name(FINAL_SOURCE.stem + "_TEMP" + FINAL_SOURCE.suffix)
CREDENTIALS   = Path(__file__).resolve().with_name("credentials.txt")

REPORT_TITLE    = "24 Hour Packet Loss Report"
REPORT_TIMEOUT  = 300        # seconds Mycom may take to compute a circle
DOWNLOAD_TIMEOUT = 300       # seconds to wait for the file after clicking Export
RETRIES         = 2          # attempts per circle
STAGGER_SECONDS = 2          # start 2nd session slightly later (gentler on login)
BACKUP_KEEP     = 48         # newest backups to keep (0 = no backups)
LOCK_MAX_AGE    = 90 * 60    # ignore a stale lock after 90 minutes
KNOWN_CIRCLES = {"BIH", "CH", "DL", "GJ", "HP", "HR", "JHK", "JK", "KK", "KL", "KO",
                 "MB", "MH", "MP", "PB", "RJ", "TN", "UE", "UW", "WB"}

try:  # never crash on non-ASCII in a Windows console
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass


def log(msg):
    print(msg, flush=True)


# ----------------------------------------------------------------------------- helpers
def read_credentials():
    if not CREDENTIALS.exists():
        raise FileNotFoundError(f"Missing credentials file: {CREDENTIALS}")
    d = {}
    for line in CREDENTIALS.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            d[k.strip().lower()] = v.strip()
    user, pwd = d.get("user") or d.get("username"), d.get("password")
    if not user or not pwd:
        raise ValueError("credentials.txt must contain user=... and password=...")
    return user, pwd


def click_first_visible(locator):
    """Click the first visible+enabled match. Returns True if clicked."""
    for i in range(locator.count()):
        item = locator.nth(i)
        try:
            if item.is_visible() and item.is_enabled():
                item.scroll_into_view_if_needed()
                item.click(timeout=8000)
                return True
        except Exception:
            pass
    return False


JS_CLICK_EXACT = """(txt) => {
  for (const el of document.querySelectorAll('*')) {
    const r = el.getBoundingClientRect(), s = getComputedStyle(el);
    if ((el.innerText || '').trim() === txt && r.width > 0 && r.height > 0 &&
        s.display !== 'none' && s.visibility !== 'hidden') {
      el.scrollIntoView({block: 'center'}); el.click(); return true;
    }
  } return false;
}"""


def until(page, fn, timeout, step_ms=400):
    """Poll fn() until truthy or timeout (seconds). Replaces fixed sleeps."""
    end = time.time() + timeout
    while time.time() < end:
        try:
            if fn():
                return True
        except Exception:
            pass
        page.wait_for_timeout(step_ms)
    return False


# ----------------------------------------------------------------------------- navigation
def open_report(context, user, pwd, circle):
    page = context.new_page()
    log(f"{circle}: opening Mycom and logging in...")
    page.goto(MYCOM_URL, wait_until="domcontentloaded", timeout=90000)
    page.locator('input[placeholder="Username"]').wait_for(state="visible", timeout=60000)
    page.locator('input[placeholder="Username"]').fill(user)
    page.locator('input[placeholder="Password"]').fill(pwd)
    page.get_by_text("Sign in", exact=True).click()

    # Web Client menu (wait for it instead of sleeping 5 s)
    web_client = page.get_by_text("Web Client", exact=True).first
    web_client.wait_for(state="visible", timeout=60000)
    web_client.click()

    bh = page.get_by_role("menuitem", name="BH", exact=True)
    if not until(page, lambda: bh.count() > 0 and bh.last.is_visible(), 20):
        raise RuntimeError("BH Web Client option was not found.")
    tabs_before = len(context.pages)
    bh.last.click()
    # BH Web Client normally opens in a new tab; if not, we stay on this one
    until(page, lambda: len(context.pages) > tabs_before, 12)
    page = context.pages[-1]
    page.wait_for_load_state("domcontentloaded")

    log(f"{circle}: selecting TWAMP...")
    twamp = page.get_by_text("TWAMP", exact=True)
    clicked, end = False, time.time() + 40
    while time.time() < end and not clicked:
        clicked = click_first_visible(twamp) or bool(page.evaluate(JS_CLICK_EXACT, "TWAMP"))
        if not clicked:
            page.wait_for_timeout(700)
    if not clicked:
        raise RuntimeError("TWAMP tab could not be selected.")

    log(f"{circle}: opening {REPORT_TITLE}...")
    tabs_before = len(context.pages)
    tile = page.get_by_text(REPORT_TITLE, exact=False)
    clicked, end = False, time.time() + 40
    while time.time() < end and not clicked:
        clicked = click_first_visible(tile)
        if not clicked:
            page.wait_for_timeout(700)
    if not clicked:
        raise RuntimeError(f"{REPORT_TITLE} tile was not found.")
    until(page, lambda: len(context.pages) > tabs_before, 10)
    page = context.pages[-1]
    page.wait_for_load_state("domcontentloaded")
    log(f"{circle}: report page opened.")
    return page


# ----------------------------------------------------------------------------- circle selection
def open_circle_list(page, selector):
    combo = selector.locator("input[role=combobox]").first
    if not combo.count():
        raise RuntimeError("Mycom Circles combobox not found.")
    combo_id = combo.get_attribute("id")
    if not combo_id:
        raise RuntimeError("Mycom Circles combobox has no id.")
    box = page.locator(f"ul#{combo_id}-list[role=listbox]")
    try:
        if box.is_visible():
            return box
    except Exception:
        pass
    arrow = selector.locator("#gwt-debug-Toolbar-button-CartDown")
    if arrow.count():
        arrow.first.click()
    else:
        combo.click()
    try:
        box.wait_for(state="visible", timeout=7000)
        return box
    except Exception:
        pass
    visible = page.locator('ul[role="listbox"]:visible')
    if visible.count() == 1:
        return visible.first
    combo.click()
    page.wait_for_timeout(800)
    visible = page.locator('ul[role="listbox"]:visible')
    if visible.count() == 1:
        return visible.first
    raise RuntimeError(f"Could not open circle list; visible listboxes={visible.count()}")


def selected(box):
    return set(box.locator('li[role="option"]').evaluate_all("""els =>
      els.filter(x => /accessibleComboOptionSelected|OptionSelected/i.test(x.className))
         .map(x => x.getAttribute('title')).filter(Boolean)"""))


def choose_circle(page, name):
    selector = page.locator("#gwt-debug-Toolbar-panel-FilterSelect")
    selector.first.wait_for(state="visible", timeout=60000)
    box = open_circle_list(page, selector)

    def toggle(circle):
        option = box.locator(f'li[role="option"][title="{circle}"]')
        if not option.count():
            raise RuntimeError(f"Circle option {circle} not found.")
        option.first.scroll_into_view_if_needed()
        option.first.click()
        page.wait_for_timeout(400)

    current = selected(box)
    log(f"{name}: current circle selection {sorted(current)}")
    if name not in current:
        toggle(name)
    for other in sorted(selected(box) - {name}):      # report opens with UW pre-selected
        log(f"{name}: deselecting {other}")
        toggle(other)
    final = selected(box)
    log(f"{name}: verified circle selection {sorted(final)}")
    if final != {name}:
        raise RuntimeError(f"Expected only {name}, got {sorted(final)}")
    apply_btn = selector.locator("#gwt-debug-Toolbar-button-Check")
    if not apply_btn.count():
        raise RuntimeError("Mycom Apply/check button not found.")
    apply_btn.first.click()
    log(f"{name}: applied; waiting for report generation...")


def wait_ready(page, name):
    """Return the Export button once Mycom has finished computing THIS circle."""
    t0 = time.time()
    saw_computing = False
    labels = page.locator("div.com-mycom-common-client-view-resource-AnalysisStyle-statusLabel")
    last_note = 0
    while time.time() - t0 < REPORT_TIMEOUT:
        computing = False
        try:
            for i in range(labels.count()):
                el = labels.nth(i)
                if el.is_visible() and re.search(r"^\s*Computing\.\.\.\s*$", el.inner_text(), re.I):
                    computing = True
                    break
        except Exception:
            pass
        saw_computing = saw_computing or computing
        # Old (UW) data can still be on screen right after Apply, so only trust Export
        # after a Computing... cycle was seen (or 30 s passed in case it was missed).
        export = page.locator("#gwt-debug-Toolbar-button-Export:visible").first
        if not computing and (saw_computing or time.time() - t0 > 30) and export.count():
            try:
                if export.is_visible() and export.is_enabled():
                    clear = export.evaluate("""el => {
                      const r = el.getBoundingClientRect();
                      const t = document.elementFromPoint(r.left + r.width/2, r.top + r.height/2);
                      return !!t && (t === el || el.contains(t));
                    }""")
                    if clear:
                        log(f"{name}: report ready after {time.time()-t0:.0f}s.")
                        return export
            except Exception:
                pass
        if time.time() - last_note > 20:
            log(f"{name}: Mycom is still generating the report ({time.time()-t0:.0f}s)...")
            last_note = time.time()
        page.wait_for_timeout(1000)
    raise TimeoutError(f"{name} report not ready within {REPORT_TIMEOUT // 60} minutes.")


# ----------------------------------------------------------------------------- download
def wait_for_download(context, downloads, timeout):
    end = time.time() + timeout
    while time.time() < end and not downloads:
        live = [p for p in context.pages if not p.is_closed()]
        if live:
            live[0].wait_for_timeout(500)       # keeps Playwright's event loop pumping
        else:
            time.sleep(0.5)
    if not downloads:
        raise RuntimeError(f"Export clicked but no download started within {timeout // 60} min.")
    return downloads[0]


def validate_download(path, name):
    from openpyxl import load_workbook
    if not path.exists() or path.stat().st_size < 5000:
        raise RuntimeError(f"{name} XLSX is missing or too small.")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if not wb.sheetnames:
            raise RuntimeError(f"{name} workbook has no worksheets.")
        found = set()
        ws = wb[wb.sheetnames[0]]
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
            if i >= 2000:
                break
            for v in row:
                if isinstance(v, str) and v.strip() in KNOWN_CIRCLES:
                    found.add(v.strip())
        log(f"{name}: sheets={wb.sheetnames}; circle codes in first rows={sorted(found)}")
        if "UW" in found and name != "UW":
            raise RuntimeError(f"{name} file contains UW rows - wrong data was exported.")
        if found - {name}:
            log(f"{name}: WARNING other circle codes present {sorted(found - {name})}")
    finally:
        wb.close()


def _download_once(name, target, user, pwd, start_delay):
    from playwright.sync_api import sync_playwright
    time.sleep(start_delay)
    pw = browser = context = page = None
    started = time.perf_counter()
    try:
        log(f"\n========== {name}: INDEPENDENT EDGE SESSION ==========")
        pw = sync_playwright().start()
        browser = pw.chromium.launch(channel="msedge", headless=False)
        context = browser.new_context(accept_downloads=True,
                                      viewport={"width": 1440, "height": 900})
        downloads = []
        context.on("page", lambda pg: pg.on("download", lambda d: downloads.append(d)))

        page = open_report(context, user, pwd, name)
        choose_circle(page, name)
        export = wait_ready(page, name)
        log(f"{name}: starting Excel export...")
        export.click()
        download = wait_for_download(context, downloads, DOWNLOAD_TIMEOUT)
        log(f"{name}: download event: {download.suggested_filename}")
        download.save_as(str(target))
        validate_download(target, name)
        log(f"{name}: saved {target} ({target.stat().st_size:,} bytes)")
        return time.perf_counter() - started
    except Exception:
        try:
            FAIL_DIR.mkdir(parents=True, exist_ok=True)
            shot = FAIL_DIR / f"{name}_{datetime.now():%Y%m%d_%H%M%S}.png"
            live = [p for p in (context.pages if context else []) if not p.is_closed()]
            (live[-1] if live else page).screenshot(path=str(shot))
            log(f"{name}: failure screenshot -> {shot}")
        except Exception:
            pass
        raise
    finally:
        for closer in (getattr(context, "close", None), getattr(browser, "close", None),
                       getattr(pw, "stop", None)):
            if closer:
                try:
                    closer()
                except Exception:
                    pass
        log(f"{name}: Edge closed.")


def download_one(name, target_str, user, pwd, start_delay):
    """Runs in its own process. Retries the whole session on failure."""
    target = Path(target_str)
    last = None
    for attempt in range(1, RETRIES + 1):
        try:
            target.unlink(missing_ok=True)
            return _download_once(name, target, user, pwd, start_delay if attempt == 1 else 0)
        except Exception as exc:
            last = exc
            log(f"{name}: attempt {attempt}/{RETRIES} failed: {exc}")
            if attempt < RETRIES:
                time.sleep(10)
    raise last


# ----------------------------------------------------------------------------- consolidation
def _trim(row):
    row = list(row)
    while row and row[-1] is None:
        row.pop()
    return row


def consolidate(files, final_path, temp_path):
    """Stream-merge files (same sheets/headers) into one workbook; returns row counts."""
    from openpyxl import Workbook, load_workbook
    log("\n========== CONSOLIDATING ==========")
    books = [load_workbook(f, read_only=True) for f in files]
    counts = {}
    try:
        names = books[0].sheetnames
        for f, b in zip(files, books):
            if b.sheetnames != names:
                raise RuntimeError(f"Worksheet mismatch in {f.name}: {b.sheetnames} vs {names}. "
                                   "Existing source will not be replaced.")
        for sheet in names:                      # cheap header check BEFORE writing anything
            first = [_trim(next(b[sheet].iter_rows(max_row=1, values_only=True), ()))
                     for b in books]
            for f, h in zip(files, first):
                if h != first[0]:
                    raise RuntimeError(f"Header mismatch in sheet {sheet} ({f.name}); "
                                       "existing source will not be replaced.")
        out = Workbook(write_only=True)
        for sheet in names:
            ws_out = out.create_sheet(sheet)
            header, per_file = None, []
            for idx, b in enumerate(books):
                n = 0
                for r_i, row in enumerate(b[sheet].iter_rows(values_only=True)):
                    if r_i == 0:
                        h = _trim(row)
                        if header is None:
                            header = h
                            ws_out.append(list(row))
                        elif h != header:
                            raise RuntimeError(f"Header mismatch in sheet {sheet} ({files[idx].name}).")
                        continue
                    if any(v is not None for v in row):
                        ws_out.append(list(row))
                        n += 1
                per_file.append(n)
            counts[sheet] = per_file
            log(f"{sheet}: " + "; ".join(f"{f.stem}={n:,}" for f, n in zip(files, per_file))
                + f"; combined={sum(per_file):,}")
        temp_path.unlink(missing_ok=True)
        out.save(temp_path)
    finally:
        for b in books:
            b.close()

    # post-save validation
    chk = load_workbook(temp_path, read_only=True)
    try:
        for sheet, per_file in counts.items():
            actual = sum(1 for row in chk[sheet].iter_rows(min_row=2, values_only=True)
                         if any(v is not None for v in row))
            if actual != sum(per_file):
                raise RuntimeError(f"Post-save validation failed for {sheet}: "
                                   f"expected {sum(per_file)}, found {actual}. Source unchanged.")
    finally:
        chk.close()
    log("Combined workbook passed validation.")

    if BACKUP_KEEP and final_path.exists():
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        backup = BACKUP_DIR / (final_path.stem + "_BACKUP_" +
                               datetime.now().strftime("%Y%m%d_%H%M%S") + final_path.suffix)
        shutil.copy2(final_path, backup)
        log(f"Previous source backed up to: {backup}")
        for old in sorted(BACKUP_DIR.glob(final_path.stem + "_BACKUP_*"))[:-BACKUP_KEEP]:
            old.unlink(missing_ok=True)

    for attempt in range(1, 7):
        try:
            os.replace(temp_path, final_path)
            break
        except PermissionError:
            log(f"{final_path.name} is locked (open in Excel/Power Query?) - retry {attempt}/6 in 10 s")
            time.sleep(10)
    else:
        raise PermissionError(f"Could not replace {final_path}. Close it and rerun. "
                              f"Combined file kept at {temp_path}")
    log(f"Combined source installed: {final_path}")

    for f in files:
        f.unlink(missing_ok=True)
    log("Temporary BIH and JHK files deleted. Only the final combined workbook remains.")


# ----------------------------------------------------------------------------- main
def acquire_lock():
    if LOCK_FILE.exists() and time.time() - LOCK_FILE.stat().st_mtime < LOCK_MAX_AGE:
        raise RuntimeError(f"Another run appears active (lock {LOCK_FILE}). Exiting.")
    LOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
    LOCK_FILE.write_text(str(os.getpid()))


def main():
    started = time.perf_counter()
    STAGING_DIR.mkdir(parents=True, exist_ok=True)
    if not FINAL_SOURCE.parent.exists():
        raise FileNotFoundError(f"Power Query source folder does not exist: {FINAL_SOURCE.parent}\n"
                                "Correct FINAL_SOURCE at the top of the script.")
    user, pwd = read_credentials()
    files = {c: STAGING_DIR / f"{c}.xlsx" for c in CIRCLES}
    for f in (*files.values(), TEMP_SOURCE):
        f.unlink(missing_ok=True)

    log("=" * 64)
    log("TWAMP " + " + ".join(CIRCLES) + " - PARALLEL EDGE DOWNLOAD")
    log(f"Power Query source: {FINAL_SOURCE}")
    log("=" * 64)

    errors = []
    ctx = multiprocessing.get_context("spawn")          # safest on Windows
    with ProcessPoolExecutor(max_workers=len(CIRCLES), mp_context=ctx) as pool:
        jobs = {pool.submit(download_one, c, str(files[c]), user, pwd, i * STAGGER_SECONDS): c
                for i, c in enumerate(CIRCLES)}
        for fut in as_completed(jobs):                  # waits until BOTH are done
            c = jobs[fut]
            try:
                log(f"{c} completed in {fut.result():.1f} seconds.")
            except Exception as exc:
                errors.append(f"{c}: {exc}")
                log(f"{c} FAILED: {exc}")

    if errors:
        raise RuntimeError("One or more downloads failed; consolidation cancelled and the "
                           "existing Power Query source was NOT replaced. " + " | ".join(errors))

    consolidate([files[c] for c in CIRCLES], FINAL_SOURCE, TEMP_SOURCE)

    h, rem = divmod(int(time.perf_counter() - started), 3600)
    m, s = divmod(rem, 60)
    log("\nSUCCESS: parallel download and consolidation completed.")
    log(f"Final combined file: {FINAL_SOURCE}")
    log(f"TOTAL PROCESSING TIME: {h:02d}:{m:02d}:{s:02d}")


if __name__ == "__main__":
    have_lock = False
    try:
        acquire_lock()
        have_lock = True
        main()
    except Exception as exc:
        log(f"\nAUTOMATION FAILED: {exc}")
        log("Existing Power Query source is not replaced on failure.")
        traceback.print_exc()
        sys.exit(1)
    finally:
        if have_lock:
            LOCK_FILE.unlink(missing_ok=True)

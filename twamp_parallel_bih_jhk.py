"""
TWAMP BIH + JHK PARALLEL DOWNLOAD AND CONSOLIDATION
Requirements: pip install playwright openpyxl
Run: python twamp_parallel_bih_jhk.py
Keep credentials.txt beside this script:
    user=YOUR_USERNAME
    password=YOUR_PASSWORD

IMPORTANT: Confirm FINAL_SOURCE matches your existing Power Query source.
The path below follows the latest screenshots.
"""
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import os, re, shutil, sys, time, traceback

MYCOM_URL = "https://welcome-page-pm-eaaprod.apps.mycom1.airtel.com/"
STAGING_DIR = Path(r"C:\TWAMP_Automation\Staging")
BACKUP_DIR = Path(r"C:\TWAMP_Automation\Backups")
FINAL_SOURCE = Path(r"E:\BOT\Power Query\Raw Dump\24 Hour Packet Loss Report_web.xlsx")
BIH_FILE, JHK_FILE = STAGING_DIR / "BIH.xlsx", STAGING_DIR / "JHK.xlsx"
TEMP_SOURCE = FINAL_SOURCE.with_name(FINAL_SOURCE.stem + "_TEMP" + FINAL_SOURCE.suffix)
CREDENTIALS_FILE = Path(__file__).resolve().with_name("credentials.txt")
REPORT_TITLE = "24 Hour Packet Loss Report"
REPORT_TIMEOUT = 300

def log(s):
    print(s, flush=True)

def credentials():
    if not CREDENTIALS_FILE.exists():
        raise FileNotFoundError(f"Missing credentials file: {CREDENTIALS_FILE}")
    d = {}
    for line in CREDENTIALS_FILE.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            d[k.strip().lower()] = v.strip()
    user, pwd = d.get("user") or d.get("username"), d.get("password")
    if not user or not pwd:
        raise ValueError("credentials.txt must contain user=... and password=...")
    return user, pwd

def open_report(context, user, pwd, circle):
    page = context.new_page()
    log(f"{circle}: opening Mycom and logging in...")
    page.goto(MYCOM_URL, wait_until="domcontentloaded", timeout=90000)
    page.locator('input[placeholder="Username"]').wait_for(state="visible", timeout=60000)
    page.locator('input[placeholder="Username"]').fill(user)
    page.locator('input[placeholder="Password"]').fill(pwd)
    page.get_by_text("Sign in", exact=True).click()
    page.wait_for_timeout(5000)

    page.get_by_text("Web Client", exact=True).first.click()
    page.wait_for_timeout(2500)
    bh = page.get_by_role("menuitem", name="BH")
    if not bh.count():
        raise RuntimeError("BH Web Client option was not found.")
    bh.last.click()
    page.wait_for_timeout(6000)
    if len(context.pages) > 1:
        page = context.pages[-1]

    log(f"{circle}: selecting TWAMP...")
    twamp = page.get_by_text("TWAMP", exact=True)
    try:
        twamp.first.wait_for(state="attached", timeout=25000)
    except Exception:
        pass
    page.wait_for_timeout(1500)
    clicked = False
    for i in range(twamp.count()):
        item = twamp.nth(i)
        try:
            if item.is_visible() and item.is_enabled():
                item.scroll_into_view_if_needed()
                item.click(timeout=10000)
                clicked = True
                break
        except Exception:
            pass
    if not clicked:
        clicked = bool(page.evaluate("""() => {
          for (const el of document.querySelectorAll('*')) {
            const r=el.getBoundingClientRect(), s=getComputedStyle(el);
            if ((el.innerText||'').trim()==='TWAMP' && r.width>0 && r.height>0 &&
                s.display!=='none' && s.visibility!=='hidden') {
              el.scrollIntoView({block:'center'}); el.click(); return true;
            }
          } return false;
        }"""))
    if not clicked:
        raise RuntimeError("TWAMP tab could not be selected.")
    page.wait_for_timeout(4000)

    before = len(context.pages)
    report = page.get_by_text(REPORT_TITLE, exact=False)
    if not report.count():
        raise RuntimeError("24 Hour Packet Loss Report tile was not found.")
    found = False
    for i in range(report.count()):
        try:
            item = report.nth(i)
            if item.is_visible():
                item.scroll_into_view_if_needed()
                item.click(timeout=12000)
                found = True
                break
        except Exception:
            pass
    if not found:
        report.first.click(timeout=15000)
    page.wait_for_timeout(4000)
    if len(context.pages) > before:
        page = context.pages[-1]
    page.wait_for_timeout(5000)
    log(f"{circle}: report page ready.")
    return page

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
         .map(x => x.getAttribute('title')).filter(Boolean)
    """))

def choose_circle(page, name):
    selector = page.locator("#gwt-debug-Toolbar-panel-FilterSelect")
    if not selector.count():
        raise RuntimeError("Mycom Circles selector not found.")
    selector.first.wait_for(state="visible", timeout=30000)
    box = open_circle_list(page, selector)

    def toggle(circle):
        option = box.locator(f'li[role="option"][title="{circle}"]')
        if not option.count():
            raise RuntimeError(f"Circle option {circle} not found.")
        option.first.scroll_into_view_if_needed()
        option.first.click()
        page.wait_for_timeout(500)

    current = selected(box)
    log(f"{name}: current circle selection {sorted(current)}")
    if name not in current:
        toggle(name)
    for other in sorted(selected(box) - {name}):
        log(f"{name}: deselecting {other}")
        toggle(other)
    final = selected(box)
    log(f"{name}: verified circle selection {sorted(final)}")
    if final != {name}:
        raise RuntimeError(f"Expected only {name}, got {sorted(final)}")
    apply = selector.locator("#gwt-debug-Toolbar-button-Check")
    if not apply.count():
        raise RuntimeError("Mycom Apply/check button not found.")
    apply.first.click()
    log(f"{name}: applied; waiting for report generation...")

def wait_ready(page, name):
    deadline = time.time() + REPORT_TIMEOUT
    while time.time() < deadline:
        computing = False
        labels = page.locator(
            "div.com-mycom-common-client-view-resource-AnalysisStyle-statusLabel"
        )
        try:
            for i in range(labels.count()):
                el = labels.nth(i)
                if el.is_visible() and re.search(r"^\\s*Computing\\.\\.\\.\\s*$", el.inner_text(), re.I):
                    computing = True
                    break
        except Exception:
            pass
        export = page.locator("#gwt-debug-Toolbar-button-Export:visible").first
        if not computing and export.count():
            try:
                if export.is_visible() and export.is_enabled():
                    clear = export.evaluate("""el => {
                      const r=el.getBoundingClientRect();
                      const t=document.elementFromPoint(r.left+r.width/2,r.top+r.height/2);
                      return !!t && (t===el || el.contains(t));
                    }""")
                    if clear:
                        log(f"{name}: report is ready.")
                        return export
            except Exception:
                pass
        log(f"{name}: Mycom is still generating the report...")
        page.wait_for_timeout(2000)
    raise TimeoutError(f"{name} report not ready within {REPORT_TIMEOUT//60} minutes.")

def validate_file(path, name):
    if not path.exists() or path.stat().st_size < 5000:
        raise RuntimeError(f"{name} XLSX is missing or too small.")
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        if not wb.sheetnames:
            raise RuntimeError(f"{name} workbook has no worksheets.")
        log(f"{name} sheets: {wb.sheetnames}")
    finally:
        wb.close()

def download_one(name, target, user, pwd):
    pw = browser = context = None
    started = time.perf_counter()
    try:
        log(f"\\n========== {name}: INDEPENDENT EDGE SESSION ==========")
        pw = sync_playwright().start()
        browser = pw.chromium.launch(channel="msedge", headless=False)
        context = browser.new_context(
            accept_downloads=True, viewport={"width": 1440, "height": 900}
        )
        page = open_report(context, user, pwd, name)
        choose_circle(page, name)
        export = wait_ready(page, name)
        log(f"{name}: starting Excel export...")
        with page.expect_download(timeout=300000) as event:
            export.click()
        download = event.value
        log(f"{name}: download event: {download.suggested_filename}")
        download.save_as(str(target))
        validate_file(target, name)
        log(f"{name}: saved {target} ({target.stat().st_size:,} bytes)")
        return time.perf_counter() - started
    finally:
        if context:
            try: context.close()
            except Exception: pass
        if browser:
            try:
                browser.close()
                log(f"{name}: Edge closed.")
            except Exception: pass
        if pw:
            try: pw.stop()
            except Exception: pass

def row_count(ws):
    return sum(1 for row in ws.iter_rows(min_row=2, values_only=True)
               if any(v is not None for v in row))

def consolidate():
    log("\\n========== CONSOLIDATING BIH + JHK ==========")
    bih = load_workbook(BIH_FILE)
    jhk = load_workbook(JHK_FILE)
    counts = {}
    try:
        if bih.sheetnames != jhk.sheetnames:
            raise RuntimeError(
                f"Worksheet mismatch. BIH={bih.sheetnames}; JHK={jhk.sheetnames}. "
                "Existing source will not be replaced."
            )
        for sheet in bih.sheetnames:
            a, b = bih[sheet], jhk[sheet]
            ah = [a.cell(1, c).value for c in range(1, a.max_column + 1)]
            bh = [b.cell(1, c).value for c in range(1, b.max_column + 1)]
            if ah != bh:
                raise RuntimeError(f"Header mismatch in sheet {sheet}; source unchanged.")
            ca, cb = row_count(a), 0
            for row in b.iter_rows(min_row=2, values_only=True):
                if any(v is not None for v in row):
                    a.append(row)
                    cb += 1
            counts[sheet] = (ca, cb)
            log(f"{sheet}: BIH={ca:,}; JHK={cb:,}; combined={ca+cb:,}")
        TEMP_SOURCE.unlink(missing_ok=True)
        bih.save(TEMP_SOURCE)
    finally:
        bih.close()
        jhk.close()

    check = load_workbook(TEMP_SOURCE, read_only=True, data_only=False)
    try:
        for sheet, (ca, cb) in counts.items():
            actual = row_count(check[sheet])
            if actual != ca + cb:
                raise RuntimeError(
                    f"Post-save validation failed for {sheet}: "
                    f"expected {ca+cb}, found {actual}. Source unchanged."
                )
    finally:
        check.close()
    log("Combined workbook passed validation.")

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    if FINAL_SOURCE.exists():
        backup = BACKUP_DIR / (
            FINAL_SOURCE.stem + "_BACKUP_" +
            datetime.now().strftime("%Y%m%d_%H%M%S") + FINAL_SOURCE.suffix
        )
        shutil.copy2(FINAL_SOURCE, backup)
        log(f"Previous source backed up to: {backup}")

    os.replace(TEMP_SOURCE, FINAL_SOURCE)
    log(f"Combined source installed: {FINAL_SOURCE}")

    for f in (BIH_FILE, JHK_FILE):
        f.unlink(missing_ok=True)
    log("Temporary BIH and JHK files deleted.")

def main():
    started = time.perf_counter()
    STAGING_DIR.mkdir(parents=True, exist_ok=True)
    if not FINAL_SOURCE.parent.exists():
        raise FileNotFoundError(
            f"Power Query source folder does not exist: {FINAL_SOURCE.parent}\\n"
            "Please correct FINAL_SOURCE at the top of the script. "
            "No folders/files will be created at a guessed location."
        )
    user, pwd = credentials()
    for f in (BIH_FILE, JHK_FILE, TEMP_SOURCE):
        f.unlink(missing_ok=True)

    log("=" * 64)
    log("TWAMP BIH + JHK - PARALLEL EDGE DOWNLOAD")
    log(f"BIH/JHK staging: {STAGING_DIR}")
    log(f"Power Query source: {FINAL_SOURCE}")
    log("=" * 64)

    errors = []
    with ThreadPoolExecutor(max_workers=2) as pool:
        jobs = {
            pool.submit(download_one, "BIH", BIH_FILE, user, pwd): "BIH",
            pool.submit(download_one, "JHK", JHK_FILE, user, pwd): "JHK",
        }
        for future in as_completed(jobs):
            circle = jobs[future]
            try:
                seconds = future.result()
                log(f"{circle} completed in {seconds:.1f} seconds.")
            except Exception as exc:
                errors.append(f"{circle}: {exc}")
                log(f"{circle} FAILED: {exc}")

    if errors:
        raise RuntimeError(
            "One or more downloads failed; consolidation cancelled and "
            "the existing Power Query source was not replaced. " + " | ".join(errors)
        )

    consolidate()
    elapsed = int(time.perf_counter() - started)
    h, rem = divmod(elapsed, 3600)
    m, s = divmod(rem, 60)
    log("\\nSUCCESS: BIH + JHK parallel download and consolidation completed.")
    log(f"Final combined file: {FINAL_SOURCE}")
    log(f"TOTAL PROCESSING TIME: {h:02d}:{m:02d}:{s:02d}")
    try:
        import winsound
        winsound.Beep(1000, 500)
    except Exception:
        print("\\a", end="", flush=True)

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log(f"\\nAUTOMATION FAILED: {exc}")
        log("Review the error above. Existing Power Query source is not replaced on failure.")
        traceback.print_exc()
        sys.exit(1)

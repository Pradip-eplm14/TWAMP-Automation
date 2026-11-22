from getpass import getpass
from pathlib import Path
import sys
import subprocess

try:
    import openpyxl
except ImportError:
    print('openpyxl is missing. Installing it now...')
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])

import re
import time
from playwright.sync_api import sync_playwright

# Number of leading JHK data rows (the rows right under the header) that are removed
# before JHK is appended to BIH. BIH is kept exactly as downloaded.
JHK_ROWS_TO_REMOVE = 5

MYCOM_URL = "https://welcome-page-pm-eaaprod.apps.mycom1.airtel.com/"
DOWNLOAD_DIR = Path(r"C:\TWAMP_Automation\Downloads")
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

from concurrent.futures import ProcessPoolExecutor, as_completed
import builtins
import multiprocessing


def download_circle_session(circle_name, destination, username, password):
    """Runs in its own process. Lines below are the working sequential logic,
    unchanged, except that it downloads only ONE circle per Edge session."""
    destination = Path(destination)

    def print(*args, **kwargs):          # prefix output so two sessions can be told apart
        kwargs.setdefault("flush", True)
        builtins.print(f"[{circle_name}]", *args, **kwargs)

    p = sync_playwright().start()
    browser = None

    try:
        print("Starting Microsoft Edge...")
        browser = p.chromium.launch(channel="msedge", headless=False)
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1440, "height": 900}
        )
        page = context.new_page()

        # 1. LOGIN
        print("Opening Mycom...")
        page.goto(MYCOM_URL, wait_until="domcontentloaded")

        print("Waiting for login page...")
        page.locator('input[placeholder="Username"]').wait_for(
            state="visible", timeout=30000
        )

        print("Entering username...")
        page.locator('input[placeholder="Username"]').fill(username)

        print("Entering password...")
        page.locator('input[placeholder="Password"]').fill(password)

        print("Clicking Sign in...")
        page.get_by_text("Sign in", exact=True).click()

        print("Waiting for Mycom welcome page...")
        page.wait_for_timeout(5000)

        # 2. WEB CLIENT -> BH
        print("Clicking Web Client...")
        page.get_by_text("Web Client", exact=True).first.click()

        print("Waiting for Web Client menu...")
        page.wait_for_timeout(2500)

        print("Selecting BH Web Client...")
        bh_menu = page.get_by_role("menuitem", name="BH")
        if bh_menu.count() == 0:
            raise RuntimeError("BH Web Client option was not found.")

        bh_menu.last.click()

        print("Waiting for BH Web Client to open...")
        page.wait_for_timeout(6000)

        if len(context.pages) > 1:
            page = context.pages[-1]
            print("Switched to the new BH Web Client tab.")

        # 3. TWAMP TAB
        # Mycom/GWT can render the homepage tabs several seconds after the
        # BH Web Client tab opens.  The older Step5_V4 / V19 working logic
        # explicitly waited for TWAMP to become attached and then used a DOM
        # fallback.  V24 had regressed to a simple count() check, which is why
        # TWAMP was intermittently reported as "not found".
        print("Selecting TWAMP tab...")
        print("Waiting for Mycom homepage tabs to finish rendering...")

        twamp = page.get_by_text("TWAMP", exact=True)

        try:
            twamp.first.wait_for(state="attached", timeout=25000)
        except Exception:
            pass

        page.wait_for_timeout(1500)
        twamp = page.get_by_text("TWAMP", exact=True)
        print(f"TWAMP text elements found: {twamp.count()}")

        clicked_twamp = False

        if twamp.count() > 0:
            for i in range(twamp.count()):
                try:
                    loc = twamp.nth(i)
                    if loc.is_visible() and loc.is_enabled():
                        loc.scroll_into_view_if_needed()
                        page.wait_for_timeout(300)
                        loc.click(timeout=10000)
                        clicked_twamp = True
                        break
                except Exception:
                    pass

            if not clicked_twamp:
                try:
                    twamp.first.click(force=True, timeout=10000)
                    clicked_twamp = True
                except Exception:
                    pass

        if not clicked_twamp:
            print("Exact TWAMP text click did not succeed; using DOM fallback...")
            try:
                result = page.evaluate("""
                    () => {
                        const els = [...document.querySelectorAll('*')];
                        for (const el of els) {
                            const r = el.getBoundingClientRect();
                            const s = getComputedStyle(el);
                            const txt = (el.innerText || '').trim();
                            if (
                                txt === 'TWAMP' &&
                                r.width > 0 &&
                                r.height > 0 &&
                                s.display !== 'none' &&
                                s.visibility !== 'hidden'
                            ) {
                                el.scrollIntoView({block:'center', inline:'center'});
                                el.click();
                                return true;
                            }
                        }
                        return false;
                    }
                """)
                clicked_twamp = bool(result)
            except Exception:
                clicked_twamp = False

        if not clicked_twamp:
            try:
                print("Visible Mycom text near TWAMP failure:")
                print(page.locator("body").inner_text(timeout=5000)[:4000])
            except Exception:
                pass
            raise RuntimeError(
                "TWAMP tab was not found after waiting 25 seconds. "
                "Mycom homepage tabs may not have finished rendering."
            )

        print("TWAMP tab selected.")
        page.wait_for_timeout(4000)

        # 4. 24 HOUR PACKET LOSS REPORT TILE
        print("Searching for 24 Hour Packet Loss Report tile...")

        report = page.get_by_text("24 Hour Packet Loss Report", exact=False)
        if report.count() == 0:
            raise RuntimeError("24 Hour Packet Loss Report tile was not found.")

        clicked = False
        for i in range(report.count()):
            try:
                loc = report.nth(i)
                if loc.is_visible():
                    print("Clicking 24 Hour Packet Loss Report tile...")
                    loc.click()
                    clicked = True
                    break
            except Exception:
                pass

        if not clicked:
            report.first.click()

        # 5. NEW REPORT TAB
        print("Waiting for the new report tab...")
        page.wait_for_timeout(4000)

        print("Looking for the newly opened 24 Hour Packet Loss Report tab...")
        report_tab = page.get_by_text("24 Hour Packet Loss Report", exact=False)

        if report_tab.count() == 0:
            raise RuntimeError(
                "The new 24 Hour Packet Loss Report tab was not found."
            )

        clicked_tab = False
        for i in range(report_tab.count()):
            try:
                loc = report_tab.nth(i)
                if loc.is_visible():
                    print("Clicking the new 24 Hour Packet Loss Report tab...")
                    loc.click()
                    clicked_tab = True
                    break
            except Exception:
                pass

        if not clicked_tab:
            report_tab.first.click()

        print("24 Hour Packet Loss Report tab selected.")
        page.wait_for_timeout(5000)

        from openpyxl import load_workbook

        circles = page.locator("#gwt-debug-Toolbar-panel-FilterSelect")
        if circles.count() == 0:
            raise RuntimeError("Mycom Circles selector was not found.")
        circles = circles.first
        circles.wait_for(state="visible", timeout=30000)

        combo = circles.locator("input[role=combobox]").first
        if combo.count() == 0:
            raise RuntimeError("Mycom Circles combobox was not found.")

        combo_id = combo.get_attribute("id")
        if not combo_id:
            raise RuntimeError("Mycom Circles combobox has no id.")

        listbox = page.locator(f"ul#{combo_id}-list[role=listbox]")

        def open_circle_list():
            nonlocal listbox

            # Keep the original Mycom dropdown behaviour: click the arrow,
            # wait for the expected listbox, then fall back to the one visible
            # listbox because Mycom/GWT can generate a different popup ID.
            try:
                if listbox.is_visible():
                    return
            except Exception:
                pass

            arrow = circles.locator("#gwt-debug-Toolbar-button-CartDown")
            if arrow.count() > 0:
                arrow.first.click()
            else:
                combo.click()

            try:
                listbox.wait_for(state="visible", timeout=7000)
                return
            except Exception:
                pass

            visible_listboxes = page.locator('ul[role="listbox"]:visible')
            count = visible_listboxes.count()

            if count == 1:
                listbox = visible_listboxes.first
                listbox.wait_for(state="visible", timeout=5000)
                print("Using Mycom visible-listbox fallback.")
                return

            # One retry via the combobox in case the arrow click did not open it.
            try:
                combo.click()
                page.wait_for_timeout(800)
            except Exception:
                pass

            visible_listboxes = page.locator('ul[role="listbox"]:visible')
            count = visible_listboxes.count()
            if count == 1:
                listbox = visible_listboxes.first
                print("Using Mycom visible-listbox fallback after retry.")
                return

            raise RuntimeError(
                "Could not open the Mycom Circles list. "
                f"Expected selector: ul#{combo_id}-list[role=listbox]. "
                f"Visible listboxes found: {count}."
            )

        def selected_circles():
            return set(
                listbox.locator('li[role="option"]').evaluate_all(
                    """els => els
                        .filter(li => /accessibleComboOptionSelected|OptionSelected/i.test(li.className))
                        .map(li => li.getAttribute('title'))
                        .filter(Boolean)"""
                )
            )

        def toggle_circle(name):
            option = listbox.locator(f'li[role="option"][title="{name}"]')
            if option.count() == 0:
                raise RuntimeError(f"Circle option '{name}' was not found.")
            option.first.scroll_into_view_if_needed()
            option.first.click()
            page.wait_for_timeout(500)

        def select_exact_circle(circle_name):
            print(f"\nSelecting only {circle_name}...")
            open_circle_list()

            current = selected_circles()
            print("Currently selected:", sorted(current))

            # First select the required circle if it is not selected.
            if circle_name not in current:
                toggle_circle(circle_name)

            # Deselect every other selected circle.
            for other in sorted(selected_circles() - {circle_name}):
                print(f"Deselecting {other}...")
                toggle_circle(other)

            final_selection = selected_circles()
            print("Verified circle selection:", sorted(final_selection))
            if final_selection != {circle_name}:
                raise RuntimeError(
                    f"Expected only {circle_name}, but Mycom shows "
                    f"{sorted(final_selection)}. Apply cancelled."
                )

            apply_button = circles.locator("#gwt-debug-Toolbar-button-Check")
            if apply_button.count() == 0:
                raise RuntimeError("Mycom Apply/check-mark button was not found.")
            apply_button.first.click()
            print(f"{circle_name} applied. Waiting for report generation...")

        def wait_for_report_ready(circle_name):
            computing = page.locator(
                "div.com-mycom-nova-common-client-view-resource-AnalysisStyle-statusLabel",
                has_text=re.compile(r"^\s*Computing\.\.\.\s*$", re.I),
            )
            deadline = time.time() + 300

            while time.time() < deadline:
                visible_computing = False
                try:
                    for i in range(computing.count()):
                        if computing.nth(i).is_visible():
                            visible_computing = True
                            break

                    if not visible_computing:
                        generic = page.get_by_text(
                            re.compile(r"^\s*Computing\.\.\.\s*$", re.I)
                        )
                        for i in range(generic.count()):
                            if generic.nth(i).is_visible():
                                visible_computing = True
                                break

                    export_button = page.locator(
                        "#gwt-debug-Toolbar-button-Export:visible"
                    ).first

                    if export_button.count() and export_button.is_visible():
                        unobstructed = page.evaluate(
                            """(el) => {
                                if (!el) return false;
                                const r = el.getBoundingClientRect();
                                const top = document.elementFromPoint(
                                    r.left + r.width / 2,
                                    r.top + r.height / 2
                                );
                                return !!top && (top === el || el.contains(top));
                            }""",
                            export_button.element_handle(),
                        )
                        if not visible_computing and unobstructed:
                            print(f"{circle_name}: report is ready.")
                            return

                except Exception:
                    pass

                print(f"{circle_name}: Mycom is still generating the report...")
                page.wait_for_timeout(2000)

            raise RuntimeError(
                f"{circle_name} report did not become ready within 5 minutes."
            )

        def download_circle_report(circle_name, destination):
            select_exact_circle(circle_name)
            wait_for_report_ready(circle_name)

            export_button = page.locator(
                "#gwt-debug-Toolbar-button-Export:visible"
            ).first
            print(f"{circle_name}: starting Excel export...")

            try:
                with page.expect_download(timeout=300000) as download_info:
                    export_button.click()
                download = download_info.value
                print(f"{circle_name} download event: {download.suggested_filename}")
                download.save_as(str(destination))
            except Exception as exc:
                raise RuntimeError(
                    f"{circle_name} Excel download failed."
                ) from exc

            # Basic ZIP/XLSX integrity check before using the download.
            check_wb = load_workbook(destination, read_only=True, data_only=False)
            try:
                if not check_wb.sheetnames:
                    raise RuntimeError(f"{circle_name} workbook has no worksheets.")
                print(f"{circle_name} workbook sheets:", check_wb.sheetnames)
            finally:
                check_wb.close()

            print(f"{circle_name} validated and saved: {destination}")
            print(f"{circle_name} file size: {destination.stat().st_size:,} bytes")

        # Download only this session's circle (BIH and JHK run in parallel sessions).
        download_circle_report(circle_name, destination)

    finally:
        # Close Edge automatically; no keyboard input is required.
        if browser:
            try:
                browser.close()
                print("Microsoft Edge closed automatically.")
            except Exception as close_error:
                print(f"Warning: Edge close reported: {close_error}")
        try:
            p.stop()
        except Exception:
            pass


def main():
    # Read Mycom credentials from credentials.txt beside this script.
    # Expected format: user=YOUR_USERNAME and password=YOUR_PASSWORD.
    CREDENTIALS_FILE = Path(__file__).with_name("credentials.txt")
    creds = {}
    if CREDENTIALS_FILE.exists():
        for line in CREDENTIALS_FILE.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            creds[key.strip().lower()] = value.strip()

    username = creds.get("user") or input("Enter Mycom username: ").strip()
    password = creds.get("password") or getpass("Enter Mycom password: ")

    # Measure the automation duration from browser startup through browser closure.
    START_TIME = time.perf_counter()

    run_succeeded = False

    try:
        print("\n========================================")
        print("TWAMP BIH + JHK AUTOMATION - PARALLEL DOWNLOAD")
        print("========================================")
        print("Required circles: BIH + JHK")
        print(f"Download folder: {DOWNLOAD_DIR}")
        print("========================================")
        print()

        from openpyxl import load_workbook
        from datetime import datetime
        import shutil
        import os

        STAGING_DIR = Path(r"C:\TWAMP_Automation\Staging")
        STAGING_DIR.mkdir(parents=True, exist_ok=True)

        BIH_FILE = STAGING_DIR / "BIH.xlsx"
        JHK_FILE = STAGING_DIR / "JHK.xlsx"

        FINAL_SOURCE = Path(
            r"E:\BOT\Power Query\Raw Dump\24 Hour Packet Loss Report_web.xlsx"
        )
        TEMP_SOURCE = FINAL_SOURCE.with_name(
            FINAL_SOURCE.stem + "_TEMP" + FINAL_SOURCE.suffix
        )

        # 6. DOWNLOAD BIH AND JHK AT THE SAME TIME (two independent Edge sessions)
        print("Starting two independent Microsoft Edge sessions: BIH and JHK...")
        jobs = {"BIH": BIH_FILE, "JHK": JHK_FILE}
        errors = []
        with ProcessPoolExecutor(
            max_workers=len(jobs),
            mp_context=multiprocessing.get_context("spawn"),
        ) as pool:
            futures = {
                pool.submit(download_circle_session, name, str(dest), username, password): name
                for name, dest in jobs.items()
            }
            for future in as_completed(futures):      # waits until BOTH are finished
                name = futures[future]
                try:
                    future.result()
                    print(f"{name} download session finished.")
                except Exception as exc:
                    errors.append(f"{name}: {exc}")
                    print(f"{name} FAILED: {exc}")
                    # full traceback from inside that Edge session, for diagnosis
                    print(str(exc.__cause__) if exc.__cause__ else "")

        if errors:
            raise RuntimeError(
                "One or more downloads failed; consolidation cancelled and the "
                "final Power Query source was NOT changed. " + " | ".join(errors)
            )

        print("\n========== BOTH DOWNLOADS COMPLETED ==========")

        # 7. CONSOLIDATE THE TWO WORKBOOKS
        print("Starting internal BIH + JHK consolidation...")

        wb_bih = load_workbook(BIH_FILE)
        wb_jhk = load_workbook(JHK_FILE)

        try:
            if wb_bih.sheetnames != wb_jhk.sheetnames:
                raise RuntimeError(
                    "BIH and JHK worksheet names/order do not match. "
                    "Final Power Query source will not be changed."
                )

            for sheet_name in wb_bih.sheetnames:
                ws_bih = wb_bih[sheet_name]
                ws_jhk = wb_jhk[sheet_name]

                bih_headers = [
                    ws_bih.cell(row=1, column=c).value
                    for c in range(1, ws_bih.max_column + 1)
                ]
                jhk_headers = [
                    ws_jhk.cell(row=1, column=c).value
                    for c in range(1, ws_jhk.max_column + 1)
                ]

                if bih_headers != jhk_headers:
                    raise RuntimeError(
                        f"Header mismatch in worksheet '{sheet_name}'. "
                        "Final Power Query source will not be changed."
                    )

                bih_count = sum(
                    1 for row in ws_bih.iter_rows(min_row=2, values_only=True)
                    if any(value is not None for value in row)
                )
                jhk_count = 0
                removed_rows = []

                for row_index, row in enumerate(ws_jhk.iter_rows(min_row=2, values_only=True)):
                    # Remove the first JHK_ROWS_TO_REMOVE rows below the header.
                    if row_index < JHK_ROWS_TO_REMOVE:
                        removed_rows.append(row)
                        continue
                    if any(value is not None for value in row):
                        ws_bih.append(row)
                        jhk_count += 1

                print(f"{sheet_name}: removed the first {len(removed_rows)} JHK "
                      f"row(s) below the header before consolidation:")
                for removed in removed_rows:
                    print("    ", [value for value in removed[:6]])
                if len(removed_rows) < JHK_ROWS_TO_REMOVE:
                    print(f"WARNING: {sheet_name} JHK has only {len(removed_rows)} "
                          f"data row(s); fewer than {JHK_ROWS_TO_REMOVE} could be removed.")

                print(f"{sheet_name}: BIH={bih_count}, "
                      f"JHK={jhk_count} (after removing {len(removed_rows)}), "
                      f"total={bih_count + jhk_count}")

            if TEMP_SOURCE.exists():
                TEMP_SOURCE.unlink()
            wb_bih.save(TEMP_SOURCE)

        finally:
            wb_bih.close()
            wb_jhk.close()

        # Validate the consolidated temporary workbook before replacing source.
        validation_wb = load_workbook(TEMP_SOURCE, read_only=True)
        try:
            for sheet_name in validation_wb.sheetnames:
                if validation_wb[sheet_name].max_row < 2:
                    raise RuntimeError(
                        f"Consolidated sheet '{sheet_name}' has no data rows."
                    )
        finally:
            validation_wb.close()

        print("Consolidated workbook validation successful.")

        # Back up the existing Power Query source, then replace it.
        if FINAL_SOURCE.exists():
            backup = FINAL_SOURCE.with_name(
                FINAL_SOURCE.stem + "_BACKUP_" +
                datetime.now().strftime("%Y%m%d_%H%M%S") +
                FINAL_SOURCE.suffix
            )
            shutil.copy2(FINAL_SOURCE, backup)
            print(f"Previous Power Query source backed up: {backup}")

        os.replace(TEMP_SOURCE, FINAL_SOURCE)

        print("\n========================================")
        print("SUCCESS: BIH AND JHK PROCESSED SEPARATELY")
        print(f"BIH file: {BIH_FILE}")
        print(f"JHK file: {JHK_FILE}")
        print(f"Final consolidated source: {FINAL_SOURCE}")
        print("Power Query source path remains unchanged.")
        print("========================================")
        run_succeeded = True

    finally:
        elapsed_seconds = int(time.perf_counter() - START_TIME)
        hours, remainder = divmod(elapsed_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        print("\n========================================")
        print(f"TOTAL PROCESSING TIME: {hours:02d}:{minutes:02d}:{seconds:02d} (HH:MM:SS)")
        print("========================================")

        # Play a short Windows beep only after a successful download and browser close.
        if run_succeeded:
            try:
                import winsound
                winsound.Beep(1000, 400)
            except Exception:
                print("\a", end="", flush=True)


if __name__ == "__main__":
    main()

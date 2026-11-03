"""Inspector 2 — see calendar HTML and form fields on create-diary-entry page."""
from playwright.sync_api import sync_playwright
import time

with sync_playwright() as p:
    browser = p.firefox.launch(headless=False, slow_mo=80)
    page = browser.new_page()
    page.goto("https://vtu.internyet.in/dashboard/student/student-diary", wait_until="networkidle")
    input("\n🔑 Log in then press ENTER...")

    time.sleep(1)

    # Step 1: Select internship via the button trigger
    print("\n--- Clicking internship button ---")
    page.get_by_role("button", name="Choose internship").click()
    time.sleep(0.8)

    # Print all options/listbox items
    print("--- Options in popover ---")
    opts = page.locator("[role='option'], [role='listbox'] *").all()
    for o in opts:
        print(f"  role={o.get_attribute('role')!r} text={o.inner_text()[:80]!r}")

    # Click first option to close
    opts[0].click() if opts else None
    time.sleep(0.5)

    # Step 2: Click date button and inspect calendar
    print("\n--- Clicking 'Pick a Date' button ---")
    page.get_by_role("button", name="Pick a Date").click()
    time.sleep(0.8)

    print("--- Calendar HTML (2000 chars) ---")
    cal_html = page.locator("[role='dialog'], [class*='calendar'], [class*='popover'], [class*='picker']").first.inner_html(timeout=3000)
    print(cal_html[:3000])

    print("\n--- All buttons in calendar ---")
    btns = page.locator("[role='dialog'] button, [class*='calendar'] button").all()
    for b in btns[:20]:
        print(f"  btn: text={b.inner_text()!r} name={b.get_attribute('name')!r} aria-label={b.get_attribute('aria-label')!r}")

    input("\nPress ENTER to close...")
    browser.close()


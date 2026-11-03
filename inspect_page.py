"""Quick inspector — opens the portal and prints all relevant element selectors."""
from playwright.sync_api import sync_playwright
import time

with sync_playwright() as p:
    browser = p.firefox.launch(headless=False, slow_mo=50)
    page = browser.new_page()

    page.goto("https://vtu.internyet.in/dashboard/student/student-diary", wait_until="networkidle")
    input("\n🔑 Log in, then press ENTER...")

    time.sleep(1)
    print("\n--- ALL INPUTS ---")
    inputs = page.locator("input").all()
    for el in inputs:
        print(f"  input: placeholder={el.get_attribute('placeholder')!r}, type={el.get_attribute('type')!r}, class={el.get_attribute('class')!r}")

    print("\n--- ALL BUTTONS ---")
    btns = page.locator("button").all()
    for el in btns:
        print(f"  button: text={el.inner_text()!r}, class={el.get_attribute('class')!r}")

    print("\n--- DIVs with 'date' in class ---")
    divs = page.locator("div[class*='date'], div[class*='Date']").all()
    for el in divs:
        print(f"  div: class={el.get_attribute('class')!r}, text={el.inner_text()[:60]!r}")

    print("\n--- SELECT elements ---")
    sels = page.locator("select").all()
    for el in sels:
        print(f"  select: class={el.get_attribute('class')!r}, name={el.get_attribute('name')!r}")

    print("\n--- Full page HTML (first 8000 chars) ---")
    html = page.content()
    print(html[:8000])

    input("\nPress ENTER to close...")
    browser.close()


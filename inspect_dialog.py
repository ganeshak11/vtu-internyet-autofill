import time
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    context = p.firefox.launch_persistent_context(
        user_data_dir="/home/ganeshak11/ScratchPad/.firefox_profile",
        headless=False,
        slow_mo=100
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.goto("https://vtu.internyet.in/dashboard/student/student-diary", wait_until="networkidle")
    
    print("If not logged in, please log in.")
    print("Then press ENTER in this terminal once you are on the student-diary page.")
    input("Waiting for user... press ENTER when ready: ")
    
    # Click Pick a Date
    print("Clicking Pick a Date...")
    page.get_by_role("button", name="Pick a Date").click()
    time.sleep(1)
    
    # Dump all elements in the dialog/popover
    popover = page.locator("[role='dialog'], [data-radix-popper-content-wrapper], div.p-3").first
    print("Popover text:\n", popover.inner_text())
    print("\nPopover HTML:\n", popover.inner_html()[:4000])
    
    context.close()


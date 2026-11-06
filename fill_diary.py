"""
VTU Internship Diary Auto-Filler
=================================
Reads diary entries from the enriched Excel file and automatically fills them into:
https://vtu.internyet.in/dashboard/student/student-diary

Usage:
    venv/bin/python3 fill_diary.py
"""

import re
import time
import openpyxl
from datetime import datetime
from playwright.sync_api import sync_playwright

# ─────────────────────────────────────────────
#  CONFIGURATION
# ─────────────────────────────────────────────
EXCEL_FILE = "INT Diary Entries_2026-27.xlsx"
FIREFOX_PROFILE_DIR = "/home/ganeshak11/ScratchPad/.firefox_profile"

FILL_FROM_DATE  = datetime(2026, 8, 17)
FILL_UNTIL_DATE = datetime(2026, 10, 31)

DELAY_BETWEEN_ENTRIES = 1.5

DEFAULT_WORK_SUMMARY = (
    "Continued working on AI/ML and Generative AI concepts as part of the internship. "
    "Studied relevant topics and practised hands-on exercises to strengthen understanding."
)
DEFAULT_LEARNINGS = (
    "Gained a better understanding of AI/ML techniques and their practical applications. "
    "Applied theoretical knowledge through exercises and improved problem-solving skills."
)

# ─────────────────────────────────────────────
#  PER-MODULE SKILL MAPPING
# ─────────────────────────────────────────────
MODULE_SKILLS = {
    "AI & ML Foundations":              ["Python", "Machine learning", "AI"],
    "Generative AI Foundations":        ["Python", "Generative ai", "AI"],
    "Generative Models":                ["Python", "Generative ai", "ML", "TensorFlow"],
    "Math Foundations":                 ["Python", "Statistical analysis", "Numpy"],
    "Deep Learning Foundations":        ["Python", "TensorFlow", "Keras", "ML"],
    "CNNs":                             ["Python", "TensorFlow", "Keras", "computer vision"],
    "Sequence Models":                  ["Python", "TensorFlow", "Keras", "Natural language processing"],
    "Deep Learning Tools":              ["Python", "TensorFlow", "PyTorch", "Keras"],
    "Deep Learning Practice":           ["Python", "TensorFlow", "Keras", "ML"],
    "Deep Learning Assessment":         ["Python", "TensorFlow", "PyTorch", "ML"],
    "VAE Foundations":                  ["Python", "TensorFlow", "Keras", "ML"],
    "VAE":                              ["Python", "TensorFlow", "Keras", "Generative ai"],
    "VAE Applications":                 ["Python", "TensorFlow", "Generative ai"],
    "VAE Security Applications":        ["Python", "TensorFlow", "ML"],
    "VAE Tools":                        ["Python", "TensorFlow", "PyTorch"],
    "VAE Assessment":                   ["Python", "TensorFlow", "ML"],
    "GAN Foundations":                  ["Python", "TensorFlow", "Keras", "Generative ai"],
    "GAN Training":                     ["Python", "TensorFlow", "Keras", "Generative ai"],
    "GANs":                             ["Python", "TensorFlow", "Keras", "Generative ai"],
    "GAN Applications":                 ["Python", "TensorFlow", "Generative ai", "computer vision"],
    "Transformers":                     ["Python", "TensorFlow", "Natural language processing", "Generative ai"],
    "Transformers Assessment":          ["Python", "TensorFlow", "Natural language processing"],
    "LLM Foundations":                  ["Python", "Natural language processing", "Generative ai", "AI"],
    "LLMs":                             ["Python", "Natural language processing", "Generative ai", "TensorFlow"],
    "Hugging Face":                     ["Python", "Natural language processing", "TensorFlow", "Generative ai"],
    "Fine-tuning LLMs":                 ["Python", "Natural language processing", "Generative ai", "PyTorch"],
    "Prompt Engineering":               ["Python", "Generative ai", "Natural language processing", "AI"],
    "RAG Foundations":                  ["Python", "Natural language processing", "Generative ai", "AI"],
    "RAG Implementation":               ["Python", "Natural language processing", "Generative ai", "MongoDB"],
    "Vector Databases":                 ["Python", "MongoDB", "Natural language processing"],
    "LangChain Basics":                 ["Python", "Natural language processing", "Generative ai"],
    "LangChain Advanced":               ["Python", "Natural language processing", "Generative ai", "MongoDB"],
    "AI Agents":                        ["Python", "Generative ai", "Natural language processing", "AI"],
    "Model Deployment":                 ["Python", "Django", "Docker"],
    "Containerization":                 ["Python", "Docker", "DevOps"],
    "MLOps Basics":                     ["Python", "DevOps", "Docker"],
    "End-to-End AI Project":            ["Python", "Django", "React", "MongoDB"],
    "Frontend Integration":             ["Python", "React", "Django"],
    "Testing AI Applications":          ["Python", "playwright Automation", "Software Testing"],
    "Performance Optimization":         ["Python", "TensorFlow", "ML"],
    "Project Documentation":            ["Python"],
    "Code Review & Refactoring":        ["Python"],
    "Project Presentation & Summary":   ["Python", "AI", "ML"],
}

def get_skills_for_module(module: str) -> list:
    module_clean = module.replace("Note:", "").strip()
    if module_clean in MODULE_SKILLS:
        skills = MODULE_SKILLS[module_clean]
    else:
        module_lower = module_clean.lower()
        skills = ["Python", "Machine learning", "AI"]
        for key, s_list in MODULE_SKILLS.items():
            if key.lower() in module_lower or module_lower in key.lower():
                skills = s_list
                break
    # Strictly ensure "Git" is never included
    return [s for s in skills if s.lower() != "git"]


# ─────────────────────────────────────────────
#  EXCEL LOADER
# ─────────────────────────────────────────────
def load_entries(excel_path: str) -> list:
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    entries = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        cols = (list(row) + [None] * 9)[:9]
        sl_no, date, hours, module, syllabus, practical, work_summary_raw, learnings_raw, _ = cols
        if date is None or not isinstance(date, datetime):
            continue
        work_summary = str(work_summary_raw or "").strip() or DEFAULT_WORK_SUMMARY
        learnings    = str(learnings_raw    or "").strip() or DEFAULT_LEARNINGS
        hours_val = 6.0
        if hours:
            h = str(hours).replace("Hrs", "").replace("hrs", "").strip()
            try:
                hours_val = float(h)
            except ValueError:
                pass
        entries.append({
            "sl_no": sl_no,
            "date": date,
            "hours": hours_val,
            "work_summary": work_summary,
            "learnings": learnings,
            "module": str(module or "").strip(),
        })
    return entries


# ─────────────────────────────────────────────
#  SKILLS SELECTION
# ─────────────────────────────────────────────
def _select_skills(page, skills: list):
    input_box = page.locator(".react-select__input-container input, input[id*='react-select']").first
    for skill in skills:
        try:
            # Click input to focus — NEVER call fill("") or backspace as that deletes previous pills!
            input_box.click()
            time.sleep(0.2)
            input_box.press_sequentially(skill, delay=25)
            time.sleep(0.35)

            # Match exact option by regex
            opt = page.locator(".react-select__option").filter(
                has_text=re.compile(rf"^\s*{re.escape(skill)}\s*$", re.IGNORECASE)
            ).first

            if opt.count() > 0:
                opt.click()
            else:
                page.keyboard.press("Enter")

            time.sleep(0.25)
            print(f"      🏷️  {skill}")
        except Exception as e:
            print(f"   ⚠️  Skill '{skill}' skipped: {e}")
            page.keyboard.press("Escape")
            time.sleep(0.1)


# ─────────────────────────────────────────────
#  FILL ONE ENTRY
# ─────────────────────────────────────────────
def fill_diary_entry(page, entry: dict):
    date: datetime = entry["date"]
    skills = get_skills_for_module(entry.get("module", ""))
    print(f"\n📅 Filling {date.strftime('%d %b %Y')}  |  {entry['module']}")

    # 1. Open diary start page
    page.goto("https://vtu.internyet.in/dashboard/student/student-diary", wait_until="networkidle")
    time.sleep(0.8)

    # 2. Select Internship dropdown
    page.locator("button:has-text('Choose internship'), button:has-text('Ethnotech')").first.click()
    time.sleep(0.3)
    page.locator("[role='option']").first.click()
    time.sleep(0.3)

    # 3. Open Date Picker
    page.get_by_role("button", name="Pick a Date").click()
    page.wait_for_selector(".rdp-months_dropdown", timeout=5000)

    # Month is 0-indexed in react-day-picker dropdown
    month_val = str(date.month - 1)
    year_val = str(date.year)
    page.select_option(".rdp-months_dropdown", value=month_val)
    page.select_option(".rdp-years_dropdown", value=year_val)
    time.sleep(0.4)

    # Click the specific day button (avoid outside month days)
    day_str = str(date.day)
    month_full = date.strftime("%B")
    clicked = False
    for b in page.locator("button.rdp-day_button").all():
        aria = b.get_attribute("aria-label") or ""
        if b.inner_text().strip() == day_str and month_full in aria:
            b.click()
            clicked = True
            break
    if not clicked:
        # Fallback to exact text match
        page.locator("button.rdp-day_button").filter(has_text=f"^{day_str}$").first.click()
    time.sleep(0.4)

    # 4. Click Continue
    page.get_by_role("button", name="Continue").click()
    page.wait_for_selector("textarea", timeout=12000)
    time.sleep(0.5)

    # 5. Fill fields
    tas = page.locator("textarea").all()
    # Work summary (ta 0)
    tas[0].click()
    tas[0].fill(entry["work_summary"])

    # Hours worked (input 0)
    h_input = page.locator("input[type='number']").first
    h_input.click()
    h_input.fill(str(int(entry["hours"])))

    # Learnings / Outcomes (ta 2)
    tas[2].click()
    tas[2].fill(entry["learnings"])

    # 6. Select Skills
    _select_skills(page, skills)
    time.sleep(0.3)

    # 7. Click Save
    page.get_by_role("button", name="Save").click()
    time.sleep(2)
    print(f"   ✅ Saved successfully!")
    time.sleep(DELAY_BETWEEN_ENTRIES)


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────
def main():
    entries = load_entries(EXCEL_FILE)
    if FILL_FROM_DATE or FILL_UNTIL_DATE:
        entries = [e for e in entries if
                   (FILL_FROM_DATE is None or e["date"] >= FILL_FROM_DATE) and
                   (FILL_UNTIL_DATE is None or e["date"] <= FILL_UNTIL_DATE)]

    print(f"📋 {len(entries)} entries to fill ({entries[0]['date'].strftime('%d %b %Y')} → {entries[-1]['date'].strftime('%d %b %Y')})")

    with sync_playwright() as p:
        context = p.firefox.launch_persistent_context(
            user_data_dir=FIREFOX_PROFILE_DIR,
            headless=False,
            slow_mo=50
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(20000)

        failed = []
        for i, entry in enumerate(entries, 1):
            print(f"\n[{i}/{len(entries)}] Progress: {i-1} done, {len(entries)-i+1} remaining")
            try:
                fill_diary_entry(page, entry)
            except Exception as e:
                print(f"   ❌ FAILED {entry['date'].strftime('%d %b %Y')}: {e}")
                failed.append(entry)
                time.sleep(2)

        print(f"\n{'='*50}")
        print(f"🎉 Completed! {len(entries)-len(failed)}/{len(entries)} entries successfully filled.")
        if failed:
            print(f"❌ Failed entries ({len(failed)}):")
            for e in failed:
                print(f"   - {e['date'].strftime('%d %b %Y')}")

        context.close()


if __name__ == "__main__":
    main()

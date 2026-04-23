#!/usr/bin/env python3
"""
Phase-05 (AU) — Agentic Scraping Orchestrator
Vendor name দিন → Claude API → scraper.py auto-generate → Demo → Confirm → Full Run → Git Push
"""

import os
import sys
import subprocess
import datetime
from pathlib import Path

import pandas as pd
import anthropic

# ── BASE PATHS ────────────────────────────────────────────────────────────────
BASE         = Path("d:/phase-05 (AU)")
AGENT        = BASE / "Agent"
VENDOR_LIST  = BASE / "Vendor List" / "SD_Web Scraping - Status Tracker.xlsx"
MEMORY_FILE  = AGENT / "Memory" / "vendor_memory.txt"
CLAUDE_INIT  = AGENT / "claude_init"
SKILLS_DIR   = AGENT / "Skills"
CHAT_DIR     = AGENT / "Chat"
CODE_DIR     = AGENT / "Code"
DEMO_DIR     = AGENT / "Demo"
DATA_DIR     = AGENT / "Data"
LOGS_FILE    = AGENT / "Logs" / "activity_log.txt"


# ── CHAT SAVER (claude_init) ──────────────────────────────────────────────────
_chat_log: list[str] = []

def chat_append(role: str, text: str):
    """Buffer a conversation turn."""
    _chat_log.append(f"[{role.upper()}] {text}")

def chat_save(vendor_name: str):
    """Save buffered conversation to Agent/claude_init/[date]_[vendor]_chat.txt"""
    CLAUDE_INIT.mkdir(parents=True, exist_ok=True)
    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    safe_name = vendor_name.replace(" ", "_")
    path = CLAUDE_INIT / f"{date_str}_{safe_name}_chat.txt"
    header = (
        f"CLAUDE INIT — {vendor_name}\n"
        f"Date : {datetime.datetime.now():%Y-%m-%d %H:%M}\n"
        f"{'='*50}\n\n"
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(header + "\n\n".join(_chat_log))
    print(f"  [claude_init] Chat saved → {path.name}")


# ── LOGGER ────────────────────────────────────────────────────────────────────
def log_activity(vendor: str, task: str, status: str, time_taken: str = "-"):
    date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    line = f"{date:<16} | {vendor:<16} | {task:<18} | {status:<12} | {time_taken}\n"
    with open(LOGS_FILE, "a", encoding="utf-8") as f:
        f.write(line)


# ── SITE DIFFICULTY CHECK ─────────────────────────────────────────────────────
DIFFICULT_VENDORS = {
    "surya":          {"difficulty": "HIGH",   "method": "Chrome Debug Port + undetected_chromedriver",
                       "vpn": False, "note": "Bot detection — Step1: debug port 9222, Step2: UC"},
    "loloi":          {"difficulty": "HIGH",   "method": "Selenium stealth",
                       "vpn": False, "note": "Heavy JS, lazy load, large catalog"},
    "liaigre":        {"difficulty": "HIGH",   "method": "undetected_chromedriver",
                       "vpn": False, "note": "Password-protected / login needed"},
    "holly hunt":     {"difficulty": "MEDIUM", "method": "Selenium stealth",
                       "vpn": False, "note": "JS-heavy, slow render"},
    "visual comfort": {"difficulty": "MEDIUM", "method": "Selenium stealth",
                       "vpn": False, "note": "Large catalog, rate limiting"},
    "janus et cie":   {"difficulty": "MEDIUM", "method": "Selenium stealth",
                       "vpn": False, "note": "Login wall for some pages"},
}

def check_difficulty(vendor_name: str) -> dict:
    key = vendor_name.lower().strip()
    for k, v in DIFFICULT_VENDORS.items():
        if k in key or key in k:
            print(f"\n  {'='*50}")
            print(f"  ⚠️  DIFFICULT SITE: {vendor_name}")
            print(f"     Difficulty : {v['difficulty']}")
            print(f"     Method     : {v['method']}")
            print(f"     Note       : {v['note']}")
            if v.get("vpn"):
                print(f"     VPN        : REQUIRED — manually connect VPN first!")
                input("     VPN ready? Press [Enter] to continue... ")
            print(f"  {'='*50}")
            return v
    return {}


# ── VENDOR SEARCH ─────────────────────────────────────────────────────────────
def search_vendor(vendor_name: str) -> dict:
    print(f"\n[1] Vendor List খুঁজছি: {vendor_name} ...")
    try:
        df = pd.ExcelFile(VENDOR_LIST).parse("Main")
        hit = df[df["Vendor"].str.contains(vendor_name, case=False, na=False)]
        if hit.empty:
            print(f"     '{vendor_name}' পাওয়া যায়নি — URL manually দিন।")
            url = input("     Website URL: ").strip()
            cat = input("     Category: ").strip()
            return {"name": vendor_name, "category": cat, "url": url, "rank": "", "remarks": ""}
        row = hit.iloc[0]
        info = {
            "name":     str(row.get("Vendor", vendor_name)),
            "category": str(row.get("Category", "")),
            "url":      str(row.get("Sample Link", "")),
            "rank":     str(row.get("Rank", "")),
            "status":   str(row.get("Scraping Status", "")),
            "remarks":  str(row.get("Remarks", "")),
        }
        print(f"     OK  → {info['name']}  |  {info['category']}  |  {info['url']}")
        return info
    except Exception as e:
        print(f"     Excel error: {e}")
        return {"name": vendor_name, "category": "", "url": "", "rank": "", "remarks": ""}


# ── MEMORY ────────────────────────────────────────────────────────────────────
def read_memory() -> str:
    print(f"\n[2] Memory পড়ছি ...")
    try:
        content = MEMORY_FILE.read_text(encoding="utf-8")
        print(f"     OK  → {len(content):,} chars loaded")
        return content
    except Exception as e:
        print(f"     Memory error: {e}")
        return ""


def read_skills() -> str:
    """Load all skill files to inject into Claude prompt."""
    if not SKILLS_DIR.exists():
        return ""
    parts = []
    for skill_file in sorted(SKILLS_DIR.glob("*.py")):
        parts.append(f"# === {skill_file.stem.upper()} ===\n"
                     + skill_file.read_text(encoding="utf-8")[:1500])
    content = "\n\n".join(parts)
    print(f"  Skills loaded: {len(list(SKILLS_DIR.glob('*.py')))} files")
    return content


def find_reference_code(memory_content: str) -> str:
    """Return first 2500 chars of an existing scraper.py as style reference."""
    for entry in memory_content.split("---"):
        for line in entry.splitlines():
            if line.strip().startswith("CODE:") or "Code Path:" in line:
                rel_path = line.split(":", 1)[1].strip()
                for fname in ["scraper.py", "step1.py", "Step1.py"]:
                    candidate = BASE / rel_path / fname
                    if candidate.exists():
                        try:
                            return candidate.read_text(encoding="utf-8")[:2500]
                        except Exception:
                            pass
    # fallback: any scraper.py in Code dir
    for p in CODE_DIR.rglob("scraper.py"):
        try:
            return p.read_text(encoding="utf-8")[:2500]
        except Exception:
            pass
    return ""


# ── CLAUDE API ────────────────────────────────────────────────────────────────
def generate_scripts(vendor_info: dict, memory_content: str, ref_code: str,
                     skills_content: str = "") -> str:
    print(f"\n[3] Claude API দিয়ে scraper.py তৈরি করছি ...")

    vendor_name = vendor_info['name']
    difficulty  = vendor_info.get('difficulty', {}).get('difficulty', 'UNKNOWN')
    anti_bot    = vendor_info.get('difficulty', {}).get('method', 'standard Selenium/requests')

    prompt = f"""You are an expert Python web-scraping engineer for a furniture/home-decor data pipeline.

====== TARGET VENDOR ======
Name       : {vendor_name}
Category   : {vendor_info['category']}
URL        : {vendor_info['url']}
Rank       : {vendor_info['rank']}
Remarks    : {vendor_info['remarks']}
Difficulty : {difficulty}
Anti-Bot   : {anti_bot}
Note       : {vendor_info.get('difficulty', {}).get('note', '')}

====== AGENT MEMORY (patterns from completed vendors) ======
{memory_content[:6000]}

====== AGENT SKILLS (helper code) ======
{skills_content[:3000]}

====== REFERENCE SCRAPER (style guide) ======
```python
{ref_code if ref_code else '# no reference available'}
```

====== TASK ======
Generate ONE complete production-ready Python script: scraper.py

### SCRAPER.PY RULES:

CONFIG (top of file):
  DEMO_MODE = True     # True = 2-5 products per category | False = all products
  BASE_URL  = "{vendor_info['url']}"
  DEMO_FILE = Path("d:/phase-05 (AU)/Agent/Demo/{vendor_name}_demo.xlsx")
  OUTPUT_FILE = Path("d:/phase-05 (AU)/Agent/Data/{vendor_name}/{vendor_name.replace(' ', '_')}.xlsx")

FLOW (one pass — no intermediate Excel):
  1. get_categories(driver) → list of (category_name, category_url)
  2. For each category:
       collect product cards → for each product → scrape_product() immediately
       If DEMO_MODE: collect 2-5 products per category, then move to next category
       If not DEMO_MODE: collect ALL products from all categories
  3. save_excel(all_rows, path)

ANTI-BOT:
  Difficulty = "{difficulty}"
  * EASY/MEDIUM  → requests+BS4 or standard Selenium headless
  * HIGH         → undetected_chromedriver (uc) OR Chrome Debug Port (port 9222)
  * Always include block detection: check for "verify you are human", "access denied", "captcha"
  * If blocked → wait 10s, retry max 3 times, then print warning

PRODUCT DETAIL EXTRACTION:
  - If Shopify: use [product_url].json API (SKU, price, description, images)
  - Extract: SKU, Description, W/D/H/Diameter/Weight, List Price
  - Convert fractions → decimals (e.g. "26 3/4" → 26.75)
  - Product Family Id = product name split on [,._-@] → take first part

EXCEL OUTPUT FORMAT (both demo and full):
  Row 1 : ["Brand", "{vendor_name}"]
  Row 2 : ["Link",  "{vendor_info['url']}"]
  Row 3 : [] (empty)
  Row 4 : column headers
  Row 5+: data rows (all categories together, sorted by category)

  Columns (fixed order):
    Index, Category, Product URL, Image URL, Product Name,
    SKU, Product Family Id, Description,
    Weight, Width, Depth, Diameter, Height,
    Seat Width, Seat Depth, Seat Height, Arm Height, Arm Width,
    List Price

  Add any DYNAMIC columns found on the site AFTER List Price.

PROGRESS: print every product → [category] [n] ProductName

AUTO-SAVE: every 20 products in full run (write partial file).
RESUME: skip products already saved if output file exists.

Output the script inside:
=== SCRAPER.PY ===
<complete code here>
"""

    chat_append("USER", prompt)

    client = anthropic.Anthropic()
    full = ""
    print("     Streaming", end="", flush=True)
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=12000,
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        for chunk in stream.text_stream:
            full += chunk
            print(".", end="", flush=True)
    print(" done")
    chat_append("CLAUDE", full)
    return full


# ── PARSE & SAVE ──────────────────────────────────────────────────────────────
def _clean_block(text: str) -> str:
    text = text.strip()
    if text.startswith("```python"):
        text = text[9:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()


def parse_and_save(response: str, vendor_name: str):
    print(f"\n[4] scraper.py সেভ করছি ...")

    out_dir = CODE_DIR / vendor_name
    out_dir.mkdir(parents=True, exist_ok=True)

    if "=== SCRAPER.PY ===" in response:
        raw = response.split("=== SCRAPER.PY ===", 1)[1]
        code = _clean_block(raw)
    else:
        code = _clean_block(response)

    scraper_path = out_dir / "scraper.py"
    scraper_path.write_text(code, encoding="utf-8")
    print(f"     Saved → {scraper_path.relative_to(BASE)}")

    _git_add(scraper_path)
    return scraper_path


# ── GIT ───────────────────────────────────────────────────────────────────────
def _git_add(path: Path):
    try:
        r = subprocess.run(["git", "add", str(path)], cwd=str(BASE),
                           capture_output=True, text=True)
        status = "staged" if r.returncode == 0 else f"err: {r.stderr.strip()}"
        print(f"     [git] {path.name} → {status}")
    except Exception as e:
        print(f"     [git] add failed: {e}")


def git_push(vendor_name: str):
    print("\n[Git] Push করছি ...")

    # Are there staged changes?
    status = subprocess.run(["git", "status", "--porcelain"], cwd=str(BASE),
                            capture_output=True, text=True).stdout
    staged = [l for l in status.splitlines() if l[:2].strip() in ("A", "M", "AM")]
    if not staged:
        print("     কোনো staged file নেই।")
        return

    # Commit
    msg = f"Add {vendor_name} scraping scripts"
    c = subprocess.run(["git", "commit", "-m", msg], cwd=str(BASE),
                       capture_output=True, text=True)
    if c.returncode != 0:
        # Already committed — amend
        subprocess.run(["git", "commit", "--amend", "--no-edit"], cwd=str(BASE))

    # Push
    p = subprocess.run(["git", "push", "origin", "main"], cwd=str(BASE),
                       capture_output=True, text=True)
    if p.returncode == 0:
        print("     GitHub-এ push হয়েছে!")
        log_activity(vendor_name, "Git Push", "Done")
    else:
        print(f"     Push ব্যর্থ: {p.stderr.strip()}")


# ── DEMO ──────────────────────────────────────────────────────────────────────
def run_demo(scraper_path: Path, vendor_name: str) -> bool:
    print(f"\n[5] Demo চালাচ্ছি (DEMO_MODE=True, 2-5 products/category) ...")
    DEMO_DIR.mkdir(parents=True, exist_ok=True)

    code = scraper_path.read_text(encoding="utf-8")
    if "DEMO_MODE" not in code:
        print("     scraper.py-এ DEMO_MODE নেই — skip।")
        return True

    # Ensure DEMO_MODE = True
    if "DEMO_MODE = False" in code:
        code = code.replace("DEMO_MODE = False", "DEMO_MODE = True", 1)
        scraper_path.write_text(code, encoding="utf-8")

    r = subprocess.run([sys.executable, str(scraper_path)],
                       cwd=str(BASE), timeout=300)
    if r.returncode == 0:
        demo_file = DEMO_DIR / f"{vendor_name}_demo.xlsx"
        if demo_file.exists():
            print(f"     Demo OK! → {demo_file.name}")
        else:
            print("     Demo ran but output file not found — check script paths.")
        return True
    else:
        print("     Demo error — script returned non-zero.")
        return False


# ── CHAT FILE ─────────────────────────────────────────────────────────────────
def vendor_from_chat() -> str:
    files = sorted(CHAT_DIR.glob("*.txt"), key=lambda f: f.stat().st_mtime, reverse=True)
    if files:
        content = files[0].read_text(encoding="utf-8").strip()
        print(f"     Chat: {files[0].name} → '{content}'")
        return content
    return ""


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 62)
    print("  PHASE-05 (AU) — AGENTIC SCRAPING ORCHESTRATOR")
    print("=" * 62)

    # ── Get vendor name ──
    if len(sys.argv) > 1:
        vendor_name = " ".join(sys.argv[1:])
    else:
        vendor_name = vendor_from_chat()
        if not vendor_name:
            vendor_name = input("\nVendor name দিন: ").strip()

    if not vendor_name:
        sys.exit("Vendor name দিতে হবে!")

    t0 = datetime.datetime.now()
    print(f"\n  Vendor : {vendor_name}")
    print(f"  Time   : {t0:%Y-%m-%d %H:%M}")

    # ── Workflow ──
    vendor_info    = search_vendor(vendor_name)
    difficulty     = check_difficulty(vendor_name)
    vendor_info["difficulty"] = difficulty
    memory_content = read_memory()
    ref_code       = find_reference_code(memory_content)
    skills_content = read_skills()
    response       = generate_scripts(vendor_info, memory_content, ref_code, skills_content)
    scraper        = parse_and_save(response, vendor_name)

    elapsed = str(datetime.datetime.now() - t0).split(".")[0]
    log_activity(vendor_name, "Script Generate", "Demo Ready", elapsed)

    # ── Demo ──
    run_demo(scraper, vendor_name)

    # ── Interactive loop ──
    print(f"\n{'='*62}")
    print("  DEMO READY — Agent/Demo/{}_demo.xlsx চেক করুন".format(vendor_name))
    print("  Commands: confirm | done | exit")
    print(f"  Code: Agent/Code/{vendor_name}/scraper.py")
    print(f"{'='*62}")

    while True:
        cmd = input("\n> ").strip().lower()

        if cmd == "confirm":
            print("\n  Full run শুরু করছি ...")
            code = scraper.read_text(encoding="utf-8")
            code = code.replace("DEMO_MODE = True", "DEMO_MODE = False", 1)
            scraper.write_text(code, encoding="utf-8")
            _git_add(scraper)

            t_run = datetime.datetime.now()
            subprocess.run([sys.executable, str(scraper)], cwd=str(BASE))

            elapsed_run = str(datetime.datetime.now() - t_run).split(".")[0]
            log_activity(vendor_name, "Full Run", "Done", elapsed_run)
            print(f"\n  Full run শেষ! ({elapsed_run})  'done' বললে push হবে।")

        elif cmd == "done":
            chat_save(vendor_name)
            git_push(vendor_name)
            print("  সব শেষ!")
            break

        elif cmd == "exit":
            chat_save(vendor_name)
            print("  বের হচ্ছি (git staged, push করা হয়নি)।")
            break

        else:
            print(f"  '{cmd}' চিনছি না। → confirm | done | exit")
            print(f"  Code edit: Agent/Code/{vendor_name}/scraper.py")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
AGENT TERMINAL - Phase-05 (AU)
Run: py -3 "d:/phase-05 (AU)/Agent/terminal.py"

Workflow:
  1. Vendor name lekho  -> code generate hobe
  2. 'demo'            -> demo run (5 products)
  3. Somossa likhle    -> Claude fix korbe + abar demo
  4. 'full run'        -> sob products
  5. 'done'            -> git push
"""

import os, sys, re, subprocess, datetime, textwrap
from pathlib import Path

import pandas as pd
import anthropic

# ── PATHS ────────────────────────────────────────────────────────────────────
BASE        = Path("d:/phase-05 (AU)")
AGENT       = BASE / "Agent"
VENDOR_XL   = AGENT / "Vendor_List" / "SD_Web Scraping - Status Tracker.xlsx"
MEMORY_FILE = AGENT / "Memory" / "agent_memory.txt"
SKILLS_DIR  = AGENT / "Skills"
CODE_DIR    = AGENT / "Code"
DEMO_DIR    = AGENT / "Demo"
DATA_DIR    = AGENT / "Data"
LOG_FILE    = AGENT / "Logs" / "activity_log.txt"
INIT_DIR    = AGENT / "claude_init"

# ── COLORS ───────────────────────────────────────────────────────────────────
RST = "\033[0m";  BLD = "\033[1m"
CYN = "\033[96m"; GRN = "\033[92m"
YLW = "\033[93m"; RED = "\033[91m"
DIM = "\033[90m"

W = 72

def hr(c="─"): print(f"{DIM}{c*W}{RST}")
def info(t):   print(f"  {YLW}{t}{RST}")
def ok(t):     print(f"  {GRN}OK  {RST}{t}")
def err(t):    print(f"  {RED}ERR {RST}{t}")
def agent(t):
    hr()
    for line in textwrap.wrap(t, W-4):
        print(f"  {CYN}{BLD}Agent>{RST} {line}")
    hr()

def prompt_user(label="You"):
    try:
        return input(f"\n{GRN}{BLD}{label} >{RST} ").strip()
    except (KeyboardInterrupt, EOFError):
        print(); return "exit"

# ── STATE ────────────────────────────────────────────────────────────────────
state = {
    "vendor":  "",
    "url":     "",
    "category": "",
    "rank":    "",
    "status":  "",
    "history": [],   # Claude chat history
}

# ── VENDOR SEARCH ────────────────────────────────────────────────────────────
def find_vendor(name: str) -> bool:
    try:
        df = pd.read_excel(VENDOR_XL)
        hit = df[df["Vendor"].str.contains(name, case=False, na=False)]
        if hit.empty:
            info(f"'{name}' Vendor List-e pawa jaini.")
            url = prompt_user("Website URL")
            cat = prompt_user("Category")
            state.update(vendor=name, url=url, category=cat, rank="?", status="New")
            return True
        r = hit.iloc[0]
        state.update(
            vendor   = str(r.get("Vendor", name)).title(),
            url      = str(r.get("Sample Link", "")),
            category = str(r.get("Category", "")),
            rank     = str(r.get("Rank", "")),
            status   = str(r.get("Scraping Status", "")),
        )
        ok(f"{state['vendor']}  |  Rank {state['rank']}  |  {state['url']}")
        return True
    except Exception as e:
        err(f"Vendor list error: {e}")
        return False

# ── MEMORY & SKILLS ──────────────────────────────────────────────────────────
def load_memory() -> str:
    try:
        return MEMORY_FILE.read_text(encoding="utf-8")
    except Exception:
        return ""

def load_skills() -> str:
    parts = []
    if SKILLS_DIR.exists():
        for f in sorted(SKILLS_DIR.glob("*.py")):
            parts.append(f"# {f.stem}\n" + f.read_text(encoding="utf-8")[:1200])
    return "\n\n".join(parts)

def load_ref_step1() -> str:
    for vendor_dir in CODE_DIR.iterdir():
        s1 = vendor_dir / "step1.py"
        if s1.exists():
            return s1.read_text(encoding="utf-8")[:2500]
    return ""

# ── GENERATE SCRIPTS (Claude API) ────────────────────────────────────────────
FIXED_COLS = (
    "Manufacturer, Source, Image URL, Product Name, SKU, Product Family Id, "
    "Description, Dimension, Weight, Width, Depth, Diameter, Length, Height"
)
DYNAMIC_COLS = (
    "then DYNAMIC columns found on the site (Seat Width, Seat Depth, Seat Height, "
    "Arm Height, Arm Width, List Price, Material, Finish, Color, Style, etc.)"
)

def generate_scripts(extra_note: str = "") -> str:
    memory  = load_memory()
    skills  = load_skills()
    ref     = load_ref_step1()
    vendor  = state["vendor"]
    url     = state["url"]
    cat     = state["category"]
    rank    = state["rank"]

    prompt = f"""You are an expert Python web-scraping engineer for a furniture/home-decor pipeline.

=== VENDOR ===
Name    : {vendor}
URL     : {url}
Category: {cat}
Rank    : {rank}
Extra   : {extra_note or 'none'}

=== AGENT MEMORY (rules + past vendors) ===
{memory[:5000]}

=== SKILLS ===
{skills[:3000]}

=== REFERENCE step1.py ===
```python
{ref}
```

=== TASK ===
Generate TWO complete Python scripts.

### STEP1.PY rules:
- DEMO_MODE = True at top (True=5 products, False=all)
- Scrape ALL categories/collections from the site
- Handle pagination (page param / infinite scroll / next button)
- Detect bot blocks: "verify you are human" / "access denied" / "captcha" -> wait 10s, retry 3x
- Shopify site? Use a[href*='/products/'] selector + collections
- Non-Shopify? Inspect and use correct selectors
- Output: d:/phase-05 (AU)/Agent/Code/{vendor}/step1_{vendor.replace(' ','_')}.xlsx
- Columns: Index, Category, Product URL, Image URL, Product Name

### STEP2.PY rules:
- Read step1 Excel (absolute path)
- Use Shopify .json API if Shopify site (product_url + '.json')
- FIXED columns in this EXACT order: {FIXED_COLS}
- After Height: add DYNAMIC columns found on site
- Product Family Id = re.split(r'[,._\\-@]', product_name)[0].strip()
- Dimension column = raw dimension string from page (e.g. "54 x 36 x 17 H")
- Convert fractions to decimals (26 3/4 -> 26.75)
- DEMO_MODE=True -> save to: d:/phase-05 (AU)/Agent/Demo/{vendor}_demo.xlsx
- DEMO_MODE=False -> save to: d:/phase-05 (AU)/Agent/Data/{vendor}/{vendor}.xlsx
- Output format:
    Row 1: Brand | {vendor}
    Row 2: Link  | {url}
    Row 3: (empty)
    Row 4: column headers
    Row 5+: data
- Auto-save every 20 products
- Print progress: [n/total] Product Name

IMPORTANT - output EXACTLY this format, no extra text:

=== STEP1.PY ===
<full step1.py code>

=== STEP2.PY ===
<full step2.py code>
"""

    state["history"].append({"role": "user", "content": prompt})

    client = anthropic.Anthropic()
    info("Claude API streaming...")
    full = ""
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        messages=state["history"],
    ) as stream:
        for chunk in stream.text_stream:
            full += chunk
            print(".", end="", flush=True)
    print()

    state["history"].append({"role": "assistant", "content": full})
    return full

# ── FIX SCRIPTS (Claude API) ─────────────────────────────────────────────────
def fix_scripts(problem: str) -> str:
    vendor = state["vendor"]
    s1 = (CODE_DIR / vendor / "step1.py").read_text(encoding="utf-8") \
         if (CODE_DIR / vendor / "step1.py").exists() else ""
    s2 = (CODE_DIR / vendor / "step2.py").read_text(encoding="utf-8") \
         if (CODE_DIR / vendor / "step2.py").exists() else ""

    prompt = f"""Fix the scraping scripts for {vendor}.

Problem reported: {problem}

Current step1.py:
```python
{s1[:3000]}
```

Current step2.py:
```python
{s2[:3000]}
```

Fix the issue and return BOTH complete scripts.
Output EXACTLY:

=== STEP1.PY ===
<fixed step1.py>

=== STEP2.PY ===
<fixed step2.py>
"""
    state["history"].append({"role": "user", "content": prompt})
    client = anthropic.Anthropic()
    info("Claude fix korche...")
    full = ""
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        messages=state["history"],
    ) as stream:
        for chunk in stream.text_stream:
            full += chunk
            print(".", end="", flush=True)
    print()
    state["history"].append({"role": "assistant", "content": full})
    return full

# ── PARSE & SAVE SCRIPTS ─────────────────────────────────────────────────────
def _clean(text: str) -> str:
    text = text.strip()
    for prefix in ["```python", "```"]:
        if text.startswith(prefix):
            text = text[len(prefix):]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()

def save_scripts(response: str) -> tuple:
    vendor = state["vendor"]
    out_dir = CODE_DIR / vendor
    out_dir.mkdir(parents=True, exist_ok=True)

    s1_code = s2_code = ""
    if "=== STEP1.PY ===" in response and "=== STEP2.PY ===" in response:
        after   = response.split("=== STEP1.PY ===", 1)[1]
        s1_raw  = after.split("=== STEP2.PY ===", 1)[0]
        s2_raw  = after.split("=== STEP2.PY ===", 1)[1]
        s1_code = _clean(s1_raw)
        s2_code = _clean(s2_raw)
    else:
        s1_code = response

    s1 = out_dir / "step1.py"
    s2 = out_dir / "step2.py"
    s1.write_text(s1_code, encoding="utf-8")
    ok(f"step1.py -> {s1}")
    if s2_code:
        s2.write_text(s2_code, encoding="utf-8")
        ok(f"step2.py -> {s2}")
    return s1, s2 if s2_code else None

# ── RUN SCRIPT ───────────────────────────────────────────────────────────────
def run_script(path: Path, demo: bool) -> bool:
    if not path or not path.exists():
        err(f"File not found: {path}")
        return False
    # Set DEMO_MODE
    code = path.read_text(encoding="utf-8")
    if demo:
        code = re.sub(r'DEMO_MODE\s*=\s*\w+', 'DEMO_MODE = True', code)
    else:
        code = re.sub(r'DEMO_MODE\s*=\s*\w+', 'DEMO_MODE = False', code)
    path.write_text(code, encoding="utf-8")

    info(f"Running: {path.name}  (DEMO_MODE={'True' if demo else 'False'})")
    r = subprocess.run([sys.executable, str(path)], cwd=str(BASE))
    return r.returncode == 0

# ── LOG ──────────────────────────────────────────────────────────────────────
def log(task, status):
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        f.write(f"{ts} | {state['vendor']:<25} | {task:<15} | {status}\n")

def save_chat():
    INIT_DIR.mkdir(parents=True, exist_ok=True)
    ts   = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
    name = state["vendor"].replace(" ", "_") or "general"
    path = INIT_DIR / f"{ts}_{name}_chat.txt"
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"TERMINAL LOG — {state['vendor']}\n")
        f.write(f"Date: {datetime.datetime.now():%Y-%m-%d %H:%M}\n{'='*60}\n\n")
        for m in state["history"]:
            f.write(f"[{m['role'].upper()}]\n{m['content'][:800]}\n\n")
    ok(f"Chat saved -> {path.name}")

# ── GIT PUSH ─────────────────────────────────────────────────────────────────
def git_push():
    vendor = state["vendor"]
    info("Git push korchi (.py files only)...")

    py_files = list((CODE_DIR / vendor).glob("*.py")) if (CODE_DIR / vendor).exists() else []
    if not py_files:
        err("Kono .py file nei.")
        return

    for f in py_files:
        subprocess.run(["git", "add", str(f)], cwd=str(BASE), capture_output=True)

    msg = f"Add {vendor} scraping scripts"
    c = subprocess.run(["git", "commit", "-m", msg], cwd=str(BASE), capture_output=True, text=True)
    p = subprocess.run(["git", "push", "origin", "master"], cwd=str(BASE), capture_output=True, text=True)
    if p.returncode == 0:
        ok("GitHub-e push hoyeche!")
        log("Git Push", "Done")
    else:
        err(f"Push fail: {p.stderr.strip()}")

# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    os.system("cls" if os.name == "nt" else "clear")
    hr("═")
    print(f"  {BLD}{CYN}PHASE-05 (AU) — AGENT TERMINAL{RST}")
    print(f"  {DIM}Commands: demo | fix | full run | done | exit{RST}")
    hr("═")

    s1 = s2 = None

    # Step 1: get vendor
    while not state["vendor"]:
        name = prompt_user("Vendor name")
        if name.lower() == "exit":
            sys.exit()
        if name:
            hr()
            info(f"Vendor List khujchi: {name}")
            if find_vendor(name):
                break

    # Step 2: generate scripts
    hr()
    info("Script generate korchi...")
    response = generate_scripts()
    s1, s2 = save_scripts(response)
    log("Generate", "Done")

    agent(
        f"Scripts ready for {state['vendor']}!\n"
        f"Code: Agent/Code/{state['vendor']}/\n\n"
        f"'demo' likhun -> demo run korbo (5 products)\n"
        f"Somossa hole describe korun -> Claude fix korbe\n"
        f"'full run' -> sob products\n"
        f"'done' -> git push"
    )

    # Step 3: interactive loop
    while True:
        cmd = prompt_user()
        if not cmd:
            continue

        cl = cmd.lower()

        # ── demo ──
        if cl in ("demo", "demo run", "demo dao", "demo deo"):
            hr()
            ok1 = run_script(s1, demo=True)
            if ok1 and s2:
                run_script(s2, demo=True)
            demo_f = DEMO_DIR / f"{state['vendor']}_demo.xlsx"
            if demo_f.exists():
                agent(f"Demo ready!\nFile: {demo_f}\n\nData thik ache? 'full run' deo.\nSomossa thakle describe koro -> Claude fix korbe.")
            else:
                agent("Demo chlecho. Agent/Demo/ folder check koro.\nSomossa thakle likhun.")
            log("Demo", "Done")

        # ── full run ──
        elif cl in ("full run", "fullrun", "confirm", "full"):
            hr()
            ok1 = run_script(s1, demo=False)
            if ok1 and s2:
                run_script(s2, demo=False)
            out = DATA_DIR / state["vendor"] / f"{state['vendor']}.xlsx"
            agent(f"Full run complete!\nFile: {out}\n'done' likhun -> git push.")
            log("Full Run", "Done")

        # ── done ──
        elif cl in ("done", "git push", "push"):
            save_chat()
            git_push()
            agent("Sob shesh! GitHub-e push hoyeche.")
            break

        # ── exit ──
        elif cl in ("exit", "quit", "ber", "bero"):
            save_chat()
            agent("Terminal bondho korchi.")
            break

        # ── fix / problem description ──
        else:
            hr()
            info(f"Problem: {cmd}")
            info("Claude fix korche...")
            response = fix_scripts(cmd)
            s1, s2 = save_scripts(response)
            agent(
                f"Fix hoyeche!\nProblem: {cmd}\n\n"
                f"'demo' deo -> abar demo run korbo."
            )
            log("Fix", cmd[:50])


if __name__ == "__main__":
    main()

#!/usr/bin/env python
"""
AGENT CHATBOX — Phase-05 (AU)
Terminal-based chat interface. Claude API দিয়ে কথা বলুন।
Commands: demo | confirm | done | exit
"""

import os
import sys
import subprocess
import datetime
from pathlib import Path

import anthropic

# ── PATHS ─────────────────────────────────────────────────────────────────────
BASE      = Path("d:/phase-05 (AU)")
AGENT     = BASE / "Agent"
CODE_DIR  = AGENT / "Code"
DEMO_DIR  = AGENT / "Demo"
DATA_DIR  = AGENT / "Data"
CHAT_INIT = AGENT / "claude_init"

# ── ANSI COLORS ───────────────────────────────────────────────────────────────
R  = "\033[0m"          # reset
B  = "\033[1m"          # bold
CY = "\033[96m"         # cyan  → Agent
GR = "\033[92m"         # green → User
YL = "\033[93m"         # yellow → system/status
RD = "\033[91m"         # red → error
DM = "\033[90m"         # dim → border

W  = os.get_terminal_size().columns if sys.stdout.isatty() else 80


def line(char="─"):
    print(f"{DM}{char * W}{R}")


def header():
    os.system("cls" if os.name == "nt" else "clear")
    line("═")
    title = "  PHASE-05 (AU) — AGENT CHATBOX"
    print(f"{B}{CY}{title}{R}")
    sub = "  Claude AI · Demo · Confirm · Done · Exit"
    print(f"{DM}{sub}{R}")
    line("═")
    print(f"{DM}  Commands:{R}  {YL}demo{R}  {YL}confirm{R}  {YL}done{R}  {YL}exit{R}  "
          f"  |  অথবা যেকোনো কথা লিখুন\n")


def print_agent(text: str):
    print(f"\n{CY}{B}Agent >{R} {text}\n")


def print_user(text: str):
    print(f"{GR}{B}You   >{R} {text}")


def print_status(text: str):
    print(f"  {YL}→ {text}{R}")


def print_error(text: str):
    print(f"  {RD}✗ {text}{R}")


# ── CHAT HISTORY ──────────────────────────────────────────────────────────────
_history: list[dict] = []
_vendor: str = ""

SYSTEM_PROMPT = """তুমি Phase-05 (AU) এর Agentic Scraping Assistant।
তুমি বাংলায় এবং সংক্ষেপে উত্তর দাও।

তোমার কাজ:
- Vendor-এর scraping script বানানো (step1.py + step2.py)
- Demo run করা (DEMO_MODE=True, 5টা product)
- Confirm হলে Full run করা
- Done বললে Git push করা

Keywords:
- "demo" → Demo run (step1 + step2, DEMO_MODE=True)
- "confirm" → Full run (DEMO_MODE=False)
- "done" → Git push
- "exit" → বের হও

সংক্ষিপ্ত ও কার্যকর উত্তর দাও।"""


def chat_with_claude(user_msg: str) -> str:
    _history.append({"role": "user", "content": user_msg})
    client = anthropic.Anthropic()
    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        system=SYSTEM_PROMPT,
        messages=_history,
    )
    reply = resp.content[0].text
    _history.append({"role": "assistant", "content": reply})
    return reply


# ── SCRIPT RUNNER ─────────────────────────────────────────────────────────────
def set_demo_mode(path: Path, mode: bool):
    code = path.read_text(encoding="utf-8")
    if mode:
        code = code.replace("DEMO_MODE = False", "DEMO_MODE = True")
    else:
        code = code.replace("DEMO_MODE = True", "DEMO_MODE = False")
    path.write_text(code, encoding="utf-8")


def run_script(script_path: Path) -> bool:
    print_status(f"Running: {script_path.name} ...")
    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=str(BASE),
    )
    return result.returncode == 0


def find_scripts(vendor: str):
    folder = CODE_DIR / vendor
    s1 = folder / "step1.py"
    s2 = folder / "step2.py"
    return (s1 if s1.exists() else None,
            s2 if s2.exists() else None)


# ── KEYWORD HANDLERS ──────────────────────────────────────────────────────────
def handle_demo():
    global _vendor
    if not _vendor:
        _vendor = input(f"  {YL}Vendor name:{R} ").strip()
    if not _vendor:
        print_error("Vendor name দিন আগে।")
        return

    s1, s2 = find_scripts(_vendor)
    if not s1:
        print_error(f"Code পাওয়া যায়নি: Agent/Code/{_vendor}/step1.py")
        print_status("আগে script তৈরি করুন অথবা vendor name ঠিক করুন।")
        return

    print_status(f"DEMO MODE — {_vendor} (5 products)")
    line()

    set_demo_mode(s1, True)
    ok1 = run_script(s1)

    if ok1 and s2:
        set_demo_mode(s2, True)
        run_script(s2)

    line()
    demo_file = DEMO_DIR / f"{_vendor.replace(' ','_')}_demo.xlsx"
    if demo_file.exists():
        print_agent(f"Demo ready!\nFile: {demo_file}\n\nConfirm করুন Full run-এর জন্য অথবা instruction দিন।")
    else:
        print_agent("Demo চলেছে। Agent/Demo/ folder চেক করুন।\nConfirm করুন অথবা instruction দিন।")


def handle_confirm():
    global _vendor
    if not _vendor:
        _vendor = input(f"  {YL}Vendor name:{R} ").strip()
    if not _vendor:
        print_error("Vendor name দিন।")
        return

    s1, s2 = find_scripts(_vendor)
    if not s1:
        print_error(f"Script পাওয়া যায়নি: Agent/Code/{_vendor}/")
        return

    print_status(f"FULL RUN — {_vendor}")
    line()

    set_demo_mode(s1, False)
    ok1 = run_script(s1)

    if ok1 and s2:
        set_demo_mode(s2, False)
        run_script(s2)

    line()
    print_agent(f"Full run শেষ! Data: Agent/Data/{_vendor}/\n'done' বললে Git push হবে।")


def handle_done():
    global _vendor
    if not _vendor:
        _vendor = input(f"  {YL}Vendor name:{R} ").strip()

    print_status("Git push করছি ...")
    vendor_label = _vendor or "scripts"
    msg = f"Add {vendor_label} scraping scripts"

    r1 = subprocess.run(["git", "add", "*.py"], cwd=str(BASE),
                        capture_output=True, text=True)
    r2 = subprocess.run(["git", "commit", "-m", msg], cwd=str(BASE),
                        capture_output=True, text=True)
    r3 = subprocess.run(["git", "push", "origin", "master"], cwd=str(BASE),
                        capture_output=True, text=True)

    if r3.returncode == 0:
        print_agent("GitHub-এ push হয়েছে!")
    else:
        print_error(f"Push ব্যর্থ: {r3.stderr.strip()}")

    # Save chat log
    save_chat_log()


def save_chat_log():
    CHAT_INIT.mkdir(parents=True, exist_ok=True)
    date_str = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
    name = _vendor.replace(" ", "_") if _vendor else "general"
    path = CHAT_INIT / f"{date_str}_{name}_chat.txt"
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"CHATBOX LOG — {_vendor or 'General'}\n")
        f.write(f"Date: {datetime.datetime.now():%Y-%m-%d %H:%M}\n")
        f.write("=" * 60 + "\n\n")
        for msg in _history:
            role = "You" if msg["role"] == "user" else "Agent"
            f.write(f"[{role}]\n{msg['content']}\n\n")
    print_status(f"Chat saved → {path.name}")


# ── MAIN LOOP ─────────────────────────────────────────────────────────────────
def main():
    global _vendor

    header()

    # Check if vendor already set via arg
    if len(sys.argv) > 1:
        _vendor = " ".join(sys.argv[1:])
        print_agent(f"Vendor: {B}{_vendor}{R}\n'demo' দিয়ে শুরু করুন।")
    else:
        print_agent("আমি Phase-05 Agent। Vendor name দিন অথবা সরাসরি কথা বলুন।")

    while True:
        try:
            user_input = input(f"{GR}{B}You   >{R} ").strip()
        except (KeyboardInterrupt, EOFError):
            print()
            save_chat_log()
            print_agent("বের হচ্ছি। আবার আসবেন!")
            break

        if not user_input:
            continue

        cmd = user_input.lower()

        # Keyword: set vendor
        if cmd.startswith("vendor:") or cmd.startswith("vendor "):
            _vendor = user_input.split(":", 1)[-1].strip() if ":" in user_input \
                      else user_input[7:].strip()
            print_agent(f"Vendor set: {B}{_vendor}{R}")

        elif cmd in ("demo", "demo run", "demo করো", "demo চালাও"):
            handle_demo()

        elif cmd in ("confirm", "full run", "confirm করো", "full run করো"):
            handle_confirm()

        elif cmd in ("done", "push", "git push", "done করো"):
            handle_done()
            break

        elif cmd in ("exit", "quit", "বের হও", "বের"):
            save_chat_log()
            print_agent("বের হচ্ছি!")
            break

        elif cmd in ("clear", "cls"):
            header()

        elif cmd == "help":
            print_agent(
                f"{YL}demo{R}     → Demo run (5 products, Excel → Demo/)\n"
                f"  {YL}confirm{R}  → Full run (সব products → Data/)\n"
                f"  {YL}done{R}     → Git push (.py files)\n"
                f"  {YL}exit{R}     → বের হও\n"
                f"  {YL}clear{R}    → Screen clear\n"
                f"  অথবা যেকোনো প্রশ্ন করুন — Claude জবাব দেবে।"
            )

        else:
            # Free chat → Claude API
            if _vendor:
                user_input = f"[Vendor: {_vendor}] {user_input}"
            try:
                reply = chat_with_claude(user_input)
                print_agent(reply)
            except Exception as e:
                print_error(f"Claude API error: {e}")


if __name__ == "__main__":
    main()

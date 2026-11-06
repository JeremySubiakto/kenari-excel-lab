#!/usr/bin/env python3
"""Cross-check the engine battery against real Excel on a Mac, with no files.

Writing a workbook into Excel's sandbox folder hung on a macOS permission prompt, and one
long script hung on an Excel dialog. So this sends each test as its own short AppleScript
to a new unsaved workbook, with alerts off. A test that hangs is recorded, Excel is
restarted by process id (only an instance this script started), and the run goes on.

This Mac's Excel uses Indonesian settings, so each formula is converted to ; and decimal commas.

Array-entered formulas fill the whole rectangle they read, so a spill shorter than the
rectangle shows #N/A where a real spill shows empty. Compare with that in mind. Real spill
behavior (H2#, #SPILL!) cannot be tested this way.

Usage: excel-crosscheck.py tests/engine-battery.json out.json
"""
import json, os, re, signal, subprocess, sys, time

battery = json.load(open(sys.argv[1]))
out_path = sys.argv[2]
START = time.time()
BUDGET = 420
PER_TEST = 14

def q(s):
    return '"' + str(s).replace('\\', '\\\\').replace('"', '\\"') + '"'

def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def osa(script, timeout):
    try:
        r = subprocess.run(["osascript", "-"], input=script, capture_output=True, text=True, timeout=timeout)
        return r.returncode, r.stdout.rstrip("\n"), r.stderr.strip()
    except subprocess.TimeoutExpired:
        return -1, "", "TIMEOUT"

def excel_pid():
    r = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True, text=True)
    return int(r.stdout.split()[0]) if r.stdout.strip() else None

if excel_pid():
    print("Excel is already open. Close it first, so this script never touches your work.")
    sys.exit(4)

def start_book():
    lines = ['tell application "Microsoft Excel"', 'set display alerts to false', 'set wb to make new workbook', 'set sh to active sheet']
    for i, row in enumerate(battery["data"]):
        for j, v in enumerate(row):
            lines.append(f'set value of range "{col(j+1)}{i+1}" of sh to {v if isinstance(v,(int,float)) else q(v)}')
    lines += ['return name of wb', 'end tell']
    rc, out, err = osa("\n".join(lines), 90)
    if rc != 0:
        raise SystemExit("Could not start Excel: " + err[:300])
    return out

def restart():
    pid = excel_pid()
    if pid:
        os.kill(pid, signal.SIGTERM)
        time.sleep(4)
        if excel_pid():
            os.kill(pid, signal.SIGKILL)
            time.sleep(2)
    return start_book()

def local(f):
    """This Mac's Excel uses Indonesian number settings: ; between arguments and , as the decimal point."""
    parts = re.split(r'("[^"]*")', f)
    for i in range(0, len(parts), 2):
        parts[i] = re.sub(r"(\d)\.(\d)", r"\1\0\2", parts[i]).replace(",", ";").replace("\0", ",")
    return "".join(parts)

def cells_of(rd):
    a, z = rd.split(":")
    m1, m2 = re.match(r"([A-Z]+)(\d+)", a), re.match(r"([A-Z]+)(\d+)", z)
    c1, c2 = col_num(m1.group(1)), col_num(m2.group(1))
    return [f"{col(c)}{r}" for r in range(int(m1.group(2)), int(m2.group(2)) + 1) for c in range(c1, c2 + 1)]

def col_num(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n

res = {}
hung = []
book = start_book()
for t in battery["tests"]:
    if time.time() - START > BUDGET:
        print("Stopped at the time budget, before", t["id"])
        break
    rd = t["read"] if ":" in t["read"] else t["read"] + ":" + t["read"]
    want = cells_of(rd)
    # A one-cell array entry is refused, so enter across two cells and read only the first.
    entry = rd if len(want) > 1 else want[0] + ":" + want[0][0] + str(int(want[0][1:]) + 1)
    reads = " & linefeed & ".join(f'(string value of range "{c}" of sh)' for c in want)
    script = "\n".join(['tell application "Microsoft Excel"', 'set sh to active sheet', f'set formula array of range "{entry}" of sh to {q(local(t["f"]))}', f'set out to {reads}', f'clear contents range "{entry}" of sh', 'return out', 'end tell'])
    rc, out, err = osa(script, PER_TEST)
    if rc == 0:
        res[t["id"]] = out.split("\n")
    elif err == "TIMEOUT":
        res[t["id"]] = ["HANG"]
        hung.append(t["id"])
        book = restart()
    else:
        res[t["id"]] = ["SCRIPT-ERROR " + err[:160]]
    json.dump(res, open(out_path, "w"), indent=0)

osa('tell application "Microsoft Excel"\nset display alerts to false\nclose active workbook saving no\nquit saving no\nend tell', 20)
print("ok", len(res), "tests, hung:", hung)

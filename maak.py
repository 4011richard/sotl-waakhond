"""Zet de codes en vrijgestelde ontvangers uit brein/outlook/waakhond.md in launchevent.js (tussen de markeringen).
Gebruik: python3 maak.py   daarna: git -C ~/.sotl-waakhond add -A && commit && push"""
import json, os, re
BREIN = os.path.expanduser("~/Step out of the line Dropbox/richard bogaerts/claude/brein/outlook/waakhond.md")
JS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "launchevent.js")
codes, vrij = ["#info"], []
for lijn in open(BREIN, encoding="utf-8"):
    x = lijn.split("#", 1)[0].strip().lstrip("-* ").replace("`", "").strip().lower() if not lijn.strip().startswith("#") else ""
    w = lijn.strip().lower()
    if w.startswith("#") and len(w) > 1 and w[1] not in "# ":
        codes.append(w.split()[0])
    elif "@" in x and " " not in x:
        vrij.append(x)
s = open(JS, encoding="utf-8").read()
s = re.sub(r"var CODES = \[.*?\];", "var CODES = " + json.dumps(sorted(set(codes))) + ";", s)
s = re.sub(r"var VRIJ = \[.*?\];", "var VRIJ = " + json.dumps(sorted(set(vrij))) + ";", s)
open(JS, "w", encoding="utf-8").write(s)
print("codes:", sorted(set(codes)), "| vrij:", sorted(set(vrij)))

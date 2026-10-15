#!/usr/bin/env python3
"""Build the single-file dashboard: src/shell.html + src/help.js + src/modules/*.html (+ src/themes/*.css) -> index.html

Each module is embedded byte-for-byte (nothing is rewritten); it is only escaped so it can live
inside a <script type="application/json"> block. To add a reconciliation, drop its HTML in
src/modules/ and add an entry to MODULES below.
"""
import json
import pathlib
import re

ROOT = pathlib.Path(__file__).parent

# id, tab label, tab subtitle, source file  (tab order = list order)
MODULES = [
    ("fastag", "FASTag", "NETC DSR × Bank",     "fastag.html"),
    ("nfs",    "NFS",    "NTSL × Bank",         "nfs.html"),
    ("rupay",  "RuPay",  "DSR × Bank",          "rupay.html"),
    ("upi",    "UPI",    "NTSL × Bank",         "upi.html"),
]


def themed(mid: str, html: str) -> str:
    """Append the dashboard theme (styling only) to the embedded copy of a tool, just before </head>.
    The files in src/modules/ are never modified; logic and rules are untouched."""
    css = (ROOT / "src" / "themes" / f"{mid}.css").read_text(encoding="utf-8")
    pat = re.compile(r"^</head>(?=\s*<body)", re.I | re.M)
    assert len(pat.findall(html)) == 1, f"{mid}: could not find the document </head>"
    block = '<style id="osr-theme">\n' + css + "</style>\n"
    return pat.sub(lambda m: block + m.group(0), html, count=1)


def embed(html: str) -> str:
    """JSON-encode and neutralise sequences that would end/alter the host <script> block."""
    s = json.dumps(html, ensure_ascii=False)
    return s.replace("</", "<\\/").replace("<!--", "<\\u0021--")


def main() -> None:
    shell = (ROOT / "src" / "shell.html").read_text(encoding="utf-8")
    mods = []
    for mid, label, desc, fname in MODULES:
        html = (ROOT / "src" / "modules" / fname).read_text(encoding="utf-8")
        mods.append({"id": mid, "label": label, "desc": desc, "html": themed(mid, html)})
    help_js = (ROOT / "src" / "help.js").read_text(encoding="utf-8")
    assert "</script" not in help_js.lower()
    assert shell.count("/*__HELP__*/") == 1
    shell = shell.replace("/*__HELP__*/", help_js)
    marker = "/*__MODULES__*/"
    assert shell.count(marker) == 1
    payload = "[" + ",".join(
        '{"id":%s,"label":%s,"desc":%s,"html":%s}'
        % (json.dumps(m["id"]), json.dumps(m["label"]), json.dumps(m["desc"]), embed(m["html"]))
        for m in mods
    ) + "]"
    out = ROOT / "index.html"
    out.write_text(shell.replace(marker, payload), encoding="utf-8")
    print(f"Built {out} ({out.stat().st_size/1024:.0f} KB, {len(mods)} modules)")


if __name__ == "__main__":
    main()

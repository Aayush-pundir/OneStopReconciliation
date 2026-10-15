#!/usr/bin/env python3
"""Build the single-file dashboard: src/shell.html + src/modules/*.html -> index.html

Each module is embedded byte-for-byte (nothing is rewritten); it is only escaped so it can live
inside a <script type="application/json"> block. To add a reconciliation, drop its HTML in
src/modules/ and add an entry to MODULES below.
"""
import json
import pathlib

ROOT = pathlib.Path(__file__).parent

# id, tab label, tab subtitle, source file  (tab order = list order)
MODULES = [
    ("fastag", "FASTag", "NETC DSR × Bank",     "fastag.html"),
    ("nfs",    "NFS",    "NTSL × Bank",         "nfs.html"),
    ("rupay",  "RuPay",  "DSR × Bank",          "rupay.html"),
    ("upi",    "UPI",    "NTSL × Bank",         "upi.html"),
]


def embed(html: str) -> str:
    """JSON-encode and neutralise sequences that would end/alter the host <script> block."""
    s = json.dumps(html, ensure_ascii=False)
    return s.replace("</", "<\\/").replace("<!--", "<\\u0021--")


def main() -> None:
    shell = (ROOT / "src" / "shell.html").read_text(encoding="utf-8")
    mods = []
    for mid, label, desc, fname in MODULES:
        html = (ROOT / "src" / "modules" / fname).read_text(encoding="utf-8")
        mods.append({"id": mid, "label": label, "desc": desc, "html": html})
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

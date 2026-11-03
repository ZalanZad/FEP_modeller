"""Extract the bid-opening table (MEKH 'Bontási adatok', 3-Nov-2026) from the PDF into inputs/bids.json.

Two names are garbled in the PDF text layer (rows 024 and 025 overlap); they are corrected here
from the readable fragments: 024 = Greenergy Wind Alpha Kft., 025 = WOC Three Wind Development Holding Kft.
"""
import json
import sys
from pathlib import Path

import pdfplumber

SRC = Path(__file__).parent / "inputs" / "Bontasi_adatok_2026-11-03.pdf"
DST = Path(__file__).parent / "inputs" / "bids.json"

NAME_FIX = {
    "024": "Greenergy Wind Alpha Kft.",
    "025": "WOC Three Wind Development Holding Korlátolt Felelősségű Társaság",
}


def num(s):
    s = (s or "").strip()
    return 0.0 if s in ("", "-") else float(s.replace(" ", "").replace(",", "."))


def main():
    with pdfplumber.open(SRC) as pdf:
        rows = pdf.pages[0].extract_tables()[0][1:]
    out = []
    for r in rows:
        r = [(c or "").replace("\n", " ").strip() for c in r]
        no, reg, sent, name, tax, node, mva, mwe, bmw, bmwh, district, site, note = r
        if no in NAME_FIX:
            name = NAME_FIX[no]
        if ";" in node:
            node_name, node_id = [p.strip() for p in node.split(";", 1)]
        else:
            node_name, node_id = "n/a", ""
        out.append({
            "no": int(no), "reg": reg, "sent": None if sent == "-" else sent, "name": name, "tax": tax,
            "node": node_name, "node_id": node_id, "mva": num(mva), "wind_mw": num(mwe),
            "bess_mw": num(bmw), "bess_mwh": num(bmwh), "district": district, "site": site, "note": note,
        })
    DST.write_text(json.dumps(out, ensure_ascii=False, indent=1))
    print(len(out), "bids written to", DST)


if __name__ == "__main__":
    main()

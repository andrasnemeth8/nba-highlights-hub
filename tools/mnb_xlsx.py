"""Robusztus MNB Power BI xlsx-olvasó.

Az MNB „Adatexportálás → Összegzett adatok" xlsx-ei két dolgot rontanak el,
amibe a szokványos olvasók (openpyxl) beletörnek:

  1. A worksheet `<dimension ref="A1"/>`-et ír, holott több oszlop van
     → openpyxl ezért csak az A oszlopot adja vissza.
  2. A `<c>` cellákon NINCS `r=` hivatkozás (csak a `<row>`-on van `r=`)
     → az oszlopokat a cellák SORRENDJÉBŐL kell visszaállítani.

Ez a modul stdlib-bel (zipfile + ElementTree) olvas, pozíció szerint, és
kezeli a `t="s"` (sharedStrings), `t="str"` (inline) és numerikus cellákat is.
Ha egy cellán mégis van `r=`, azt tiszteletben tartja (rés-tűrő).
"""
from __future__ import annotations
import re
import zipfile
from xml.etree import ElementTree as ET

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_COL_RE = re.compile(r"[A-Z]+")


def _col_to_idx(ref: str) -> int:
    """'B3' -> 1 (0-alapú oszlopindex)."""
    letters = _COL_RE.match(ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _num(txt: str):
    """Számszöveg → int, ha egész, különben float; ha nem szám, marad szöveg."""
    if txt is None:
        return None
    try:
        if "." not in txt and "e" not in txt.lower():
            return int(txt)
        return float(txt)
    except ValueError:
        return txt


def read_sheet(path: str, sheet: str = "xl/worksheets/sheet1.xml") -> list[list]:
    """A munkalapot sorok listájaként adja vissza (minden sor cellák listája).

    Üres cella -> None. A számcellák int/float-ként jönnek vissza.
    """
    with zipfile.ZipFile(path) as z:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sroot.findall(f"{_NS}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{_NS}t")))
        root = ET.fromstring(z.read(sheet))

    data = root.find(f"{_NS}sheetData")
    rows: list[list] = []
    for row in data.findall(f"{_NS}row"):
        cells: list = []
        col = 0
        for c in row.findall(f"{_NS}c"):
            ref = c.get("r")
            if ref:                       # ha van hivatkozás, az a mérvadó (rés-tűrés)
                col = _col_to_idx(ref)
            t = c.get("t")
            v = c.find(f"{_NS}v")
            if t == "s":                  # sharedStrings index
                val = shared[int(v.text)] if v is not None and v.text is not None else None
            elif t == "str":              # inline képlet-eredmény / szöveg
                val = v.text if v is not None else None
            elif t == "inlineStr":
                inl = c.find(f"{_NS}is")
                val = "".join(x.text or "" for x in inl.iter(f"{_NS}t")) if inl is not None else None
            elif v is not None:           # numerikus
                val = _num(v.text)
            else:                         # üres cella
                val = None
            while len(cells) <= col:
                cells.append(None)
            cells[col] = val
            col += 1
        rows.append(cells)
    return rows


def read_table(path: str, header_marker: str | None = None):
    """Az MNB-export 'táblázatos' részét adja vissza: (header, data_rows).

    Az MNB-export elején van pár 'Alkalmazott szűrők:' / üres sor; az igazi
    fejléc az első olyan sor, ami >1 nem üres cellát tartalmaz (vagy amelyik
    `header_marker`-rel kezdődik). Onnantól jönnek az adatsorok.
    """
    rows = read_sheet(path)
    start = None
    for i, r in enumerate(rows):
        nonempty = [c for c in r if c not in (None, "")]
        if header_marker is not None:
            if r and r[0] == header_marker:
                start = i
                break
        elif len(nonempty) > 1:
            start = i
            break
    if start is None:
        raise ValueError(f"Nem találom a fejlécsort: {path}")
    header = rows[start]
    data = [r for r in rows[start + 1:] if any(c not in (None, "") for c in r)]
    return header, data


if __name__ == "__main__":
    import sys
    h, d = read_table(sys.argv[1])
    print("HEADER:", h)
    print("ROWS:", len(d))
    for r in d[:5]:
        print(r)

"""MNB befektetési-alap xlsx exportok -> REPORT_DATA JSON -> beágyazás a riportba.

Ez váltja ki a korábbi ad-hoc PowerShell-eljárást. Lépések:
  1. Beolvassa az összes MNB xlsx-et a forrásmappából a robusztus, pozíciós
     olvasóval (mnb_xlsx.read_table) — kezeli az `r=` nélküli cellákat.
  2. A fájlokat NEM fájlnév, hanem FEJLÉC alapján azonosítja (a Power BI
     duplikált letöltései — (1)/(2) — így is helyre kerülnek), és minden
     szekcióhoz a legtöbb hónapot lefedő fájlt választja.
  3. Összeállítja a REPORT_DATA objektumot (flows/assets/liabs/owners/ownersT/geo).
  4. --check : összeveti a riportba már beágyazott adattal (regresszió).
     --write : beinjektálja az új adatot a HTML-be.

Használat:
  python tools/build_report.py --check
  python tools/build_report.py --write
  python tools/build_report.py --src "C:\\...\\Downloads" --html "...\\befektetesi-alapok-riport.html" --write
"""
from __future__ import annotations
import argparse
import glob
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mnb_xlsx import read_table

HU_MONTHS = ["Január", "Február", "Március", "Április", "Május", "Június",
             "Július", "Augusztus", "Szeptember", "Október", "November", "December"]
HMON = {m: i + 1 for i, m in enumerate(HU_MONTHS)}

# A típusok kanonikus sorrendje (a meglévő riport szerint); ismeretlen típus a végére kerül.
CANON_TYPES = ["Pénzpiaci", "Kötvény", "Részvény", "Kockázati- és magántőke", "Vegyes",
               "Garantált", "Származtatott", "Ingatlan", "Árupiaci", "Egyéb"]

OWNER_SECTORS = ["Háztartások", "Pénzügyi vállalatok", "Nem pénzügyi vállalatok", "Államháztartás"]

# A megjelenítendő oszlopfeliratok KÉZZEL RÖVIDÍTETT változatai (a nyers MNB-fejlécek
# túl hosszúak a chartokhoz). Ezek a riportba szántak — nem a fájl fejlécéből jönnek.
FLOW_COLS = ["Nyitó", "Árváltozás", "Devizaárfolyam", "Tranzakció",
             "Egyéb volumen", "Változás össz.", "Záró", "Darabszám"]
ASSET_COLS = ["Készpénz és betétek", "Hitelviszonyt megtestesítő ép.", "Részvények és részesedések",
              "Befektetési jegyek", "Derivatívák", "Nem pénzügyi eszközök", "Egyéb eszközök", "Eszközök összesen"]
LIAB_COLS = ["Kibocsátott befektetési jegyek (NEÉ)", "Felvett hitelek",
             "Derivatívák (forrás oldal)", "Egyéb források", "Források összesen"]

META_SOURCE = "MNB – Befektetési alapok statisztika"

# Szekció-felismerés a fejléc egy jellemző cellája alapján (magyar export!).
SECTION_MARKERS = {
    "flows":  "Nyitó állomány",
    "assets": "Készpénz és betétek",
    "liabs":  "Kibocsátott befektetési jegyek (nettó eszközérték)",
    "owners": "Háztartások és háztartásokat segítő nonprofit intézmények",
}


def month_key(m: str):
    """'2026. Április' -> (2026, 4) rendezéshez."""
    year = int(m[:4])
    name = m.split(" ", 1)[1].strip()
    return (year, HMON[name])


def clean(x):
    """3 tizedesre kerekít; egész értéket int-ként ad vissza (a JSON tömörebb)."""
    if x is None:
        return 0
    f = round(float(x), 3)
    return int(f) if f == int(f) else f


def classify(header: list) -> str | None:
    cells = [str(c).strip() for c in header if c is not None]
    if cells[:1] == ["Ország"] and any("Értékpapír" in c for c in cells):
        return "geo"
    for sect, marker in SECTION_MARKERS.items():
        if marker in cells:
            return sect
    return None  # pl. angol export -> kihagyjuk


def scan_sources(src_dir: str):
    """Minden xlsx-et beolvas, fejléc alapján osztályoz, szekciónként a
    legtöbb hónapot lefedő (geo-nál legtöbb soros) fájlt választja."""
    by_section = {}  # section -> list of (n_metric, header, data, path)
    for p in sorted(glob.glob(os.path.join(src_dir, "*.xlsx"))):
        try:
            header, data = read_table(p)
        except Exception:
            continue
        sect = classify(header)
        if not sect:
            continue
        if sect == "geo":
            score = len(data)
        else:
            score = len({r[0] for r in data})  # különböző hónapok száma
        by_section.setdefault(sect, []).append((score, header, data, os.path.basename(p)))

    chosen = {}
    for sect, cands in by_section.items():
        cands.sort(key=lambda t: t[0], reverse=True)
        chosen[sect] = cands[0]
        if len(cands) > 1:
            others = ", ".join(f"{c[3]}({c[0]})" for c in cands[1:])
            print(f"  [{sect}] kiválasztva: {cands[0][3]} ({cands[0][0]})  | mellőzve: {others}")
        else:
            print(f"  [{sect}] kiválasztva: {cands[0][3]} ({cands[0][0]})")
    return chosen


def collect_months_types(chosen):
    months, types = set(), set()
    for sect in ("flows", "assets", "liabs", "owners"):
        if sect not in chosen:
            continue
        _, _, data, _ = chosen[sect]
        for r in data:
            months.add(r[0])
            types.add(r[1])
    months = sorted(months, key=month_key)
    type_order = [t for t in CANON_TYPES if t in types] + [t for t in sorted(types) if t not in CANON_TYPES]
    return months, type_order


def build_grid(data, months, types, ncols, col_start=2):
    """type -> [hónaponként ncols érték]; hiányzó (típus,hónap) -> nullák."""
    midx = {m: i for i, m in enumerate(months)}
    grid = {t: [[0] * ncols for _ in months] for t in types}
    for r in data:
        m, t = r[0], r[1]
        if t not in grid or m not in midx:
            continue
        vals = [clean(v) for v in r[col_start:col_start + ncols]]
        vals += [0] * (ncols - len(vals))
        grid[t][midx[m]] = vals[:ncols]
    return grid


def build_report(src_dir: str):
    print("Forrásfájlok kiválasztása fejléc alapján:")
    chosen = scan_sources(src_dir)
    missing = [s for s in ("flows", "assets", "liabs", "owners") if s not in chosen]
    if missing:
        print(f"  FIGYELEM: hiányzó szekciók: {missing}")

    months, types = collect_months_types(chosen)
    print(f"\nHónapok: {len(months)} ({months[0]} – {months[-1]}), típusok: {len(types)}")

    D = {
        "meta": {
            "updated": months[-1],
            "unit": "Mrd Ft",
            "source": META_SOURCE,
            "range": f"{months[0]} – {months[-1]}",
        },
        "types": types,
        "months": months,
    }

    # oszlopfeliratok: a gondozott, rövidített címkék (lásd fent), nem a nyers fejléc
    if "flows" in chosen:
        D["flowCols"] = FLOW_COLS
        D["flows"] = build_grid(chosen["flows"][2], months, types, 8)
    if "assets" in chosen:
        D["assetCols"] = ASSET_COLS
        D["assets"] = build_grid(chosen["assets"][2], months, types, 8)
    if "liabs" in chosen:
        D["liabCols"] = LIAB_COLS
        D["liabs"] = build_grid(chosen["liabs"][2], months, types, 5)
    if "owners" in chosen:
        D["ownerSectors"] = OWNER_SECTORS
        ownersT = build_grid(chosen["owners"][2], months, types, 4)  # 4 szektoroszlop
        D["ownersT"] = ownersT
        # havi összesen = típusok összege szektoronként
        owners = []
        for i in range(len(months)):
            owners.append([clean(sum(ownersT[t][i][s] for t in types)) for s in range(4)])
        D["owners"] = owners

    if "geo" in chosen:
        _, _, gdata, _ = chosen["geo"]
        D["geo"] = [[r[0], r[1], clean(r[2])] for r in gdata if r[0] is not None]

    return D


# ---------------- validáció a beágyazott adattal ----------------

def load_embedded(html_path: str):
    html = open(html_path, encoding="utf-8").read()
    m = re.search(r"REPORT_DATA\s*=\s*(\{.*?\})\s*;?\s*</script>", html, re.S)
    return json.loads(m.group(1)) if m else None


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def diff_section(name, emb, new, tol=0.0015):
    """Két (lehet beágyazott) struktúra numerikus eltérése."""
    total = mism = 0
    examples = []

    def walk(a, b, path=""):
        nonlocal total, mism
        if isinstance(a, dict):
            for k in a:
                walk(a[k], (b or {}).get(k), f"{path}.{k}")
        elif isinstance(a, list):
            for i, av in enumerate(a):
                walk(av, b[i] if isinstance(b, list) and i < len(b) else None, f"{path}[{i}]")
        else:
            na, nb = _num(a), _num(b)
            if na is not None:
                total += 1
                if nb is None or abs(na - nb) > tol:
                    mism += 1
                    if len(examples) < 6:
                        examples.append(f"{path}: beágyazott={a} új={b}")
            elif a != b:
                total += 1
                mism += 1
                if len(examples) < 6:
                    examples.append(f"{path}: beágyazott={a!r} új={b!r}")

    walk(emb, new, name)
    status = "OK ✓" if mism == 0 else f"ELTÉRÉS ({mism}/{total})"
    print(f"  {name:12s} {status}")
    for e in examples:
        print("      X", e)
    return mism


def check(D, html_path):
    emb = load_embedded(html_path)
    if not emb:
        print("Nincs beágyazott REPORT_DATA a HTML-ben.")
        return
    print("\n=== VALIDÁCIÓ a beágyazott adattal ===")
    total_mism = 0
    for key in ("types", "months", "flowCols", "assetCols", "liabCols", "ownerSectors",
                "flows", "assets", "liabs", "owners", "ownersT", "geo"):
        if key in emb or key in D:
            total_mism += diff_section(key, emb.get(key), D.get(key))
    print(f"\nÖSSZESEN eltérés: {total_mism}")


def write_html(D, html_path):
    html = open(html_path, encoding="utf-8").read()
    payload = json.dumps(D, ensure_ascii=False, separators=(",", ":"))
    new = re.sub(r"(REPORT_DATA\s*=\s*)\{.*?\}(\s*;?\s*</script>)",
                 lambda m: m.group(1) + payload + m.group(2), html, count=1, flags=re.S)
    if new == html:
        print("FIGYELEM: nem találtam a REPORT_DATA blokkot — nincs írás.")
        return
    open(html_path, "w", encoding="utf-8").write(new)
    print(f"Beírva: {html_path}  ({len(payload)} byte adat)")


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    default_html = os.path.join(os.path.dirname(here), "befektetesi-alapok-riport.html")
    default_src = os.path.join(os.environ.get("USERPROFILE", ""), "Downloads")

    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=default_src, help="forrásmappa az xlsx-ekkel")
    ap.add_argument("--html", default=default_html, help="cél HTML-riport")
    ap.add_argument("--check", action="store_true", help="összevetés a beágyazott adattal")
    ap.add_argument("--write", action="store_true", help="beírás a HTML-be")
    ap.add_argument("--dump", help="REPORT_DATA kiírása JSON-fájlba")
    args = ap.parse_args()

    D = build_report(args.src)
    if args.dump:
        json.dump(D, open(args.dump, "w", encoding="utf-8"), ensure_ascii=False)
        print(f"\nJSON kiírva: {args.dump}")
    if args.check or not args.write:
        check(D, args.html)
    if args.write:
        write_html(D, args.html)


if __name__ == "__main__":
    main()

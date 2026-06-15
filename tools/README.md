# MNB riport-konverter (Python)

Az MNB *Befektetési alapok statisztikai mérlege* Power BI-riportból exportált
`.xlsx` fájlokat alakítja át a HTML-riportokba ágyazott `REPORT_DATA` JSON-objektummá.
Ez váltja le a korábbi ad-hoc PowerShell-eljárást.

## Fájlok

| fájl | szerep |
|------|--------|
| [`mnb_xlsx.py`](mnb_xlsx.py) | Robusztus, **pozíciós** xlsx-olvasó (csak stdlib: `zipfile` + `xml`) |
| [`build_report.py`](build_report.py) | Forrásfájlok → `REPORT_DATA` összeállítás, validáció, beírás a HTML-be |

## Miért nem elég az openpyxl?

Az MNB-export két dolgot rosszul ír ki, amibe a szokványos olvasók beletörnek:

1. `<dimension ref="A1"/>` — pedig sok oszlop van → az openpyxl csak az **A** oszlopot adja vissza.
2. A `<c>` cellákon **nincs `r=` hivatkozás** (csak a `<row>`-on) → az oszlopokat a cellák
   **sorrendjéből** kell visszaállítani.

Ezért az `mnb_xlsx.read_table()` saját, pozíciós olvasással dolgozik (és ha mégis van
`r=`, azt tiszteletben tartja, így rés-tűrő).

## Forrás → szekció leképezés

A szkript **a fájl fejléce alapján** azonosítja a szerepet (nem a fájlnévből — a Power BI
duplikált letöltései, `(1)`/`(2)`, így is helyre kerülnek), és szekciónként a **legtöbb
hónapot** lefedő fájlt választja.

| szekció | felismerő fejléc-cella | oszlopok |
|--------|------------------------|----------|
| `flows` | „Nyitó állomány" | 8 (Nyitó, Árváltozás, Deviza, Tranzakció, Egyéb, Össz, Záró, Darabszám) |
| `assets` | „Készpénz és betétek" | 8 (eszközösszetétel + összesen) |
| `liabs` | „Kibocsátott befektetési jegyek (nettó eszközérték)" | 5 (forrásoldal) |
| `ownersT`/`owners` | „Háztartások és háztartásokat segítő nonprofit intézmények" | 4 szektor |
| `geo` | `Ország` + `Értékpapír` | 3 (ország, instrumentum, érték) |

A megjelenített oszlopcímkék kézzel rövidített változatai a szkript tetején vannak
(`FLOW_COLS`, `ASSET_COLS`, `LIAB_COLS`) — szándékosan nem a nyers (hosszú) MNB-fejlécek.

## Eljárás: adatfrissítés új hónappal

1. Az MNB Power BI-riportban (sta.mnb.hu) minden érintett vizuálnál:
   **… → Adatexportálás → Összegzett adatok → .xlsx**. A fájlok a `Downloads`-ba kerülnek.
   Kellenek: idősoros (alaptípus-bontás), eszközök, források, tulajdonosi szektor-bontás,
   földrajzi eloszlás. (A régi fájlokat nem kell törölni — a szkript a legfrissebbet választja.)
2. Ellenőrzés a jelenlegi beágyazott adattal:
   ```
   python tools/build_report.py --check
   ```
3. Beírás a riportokba:
   ```
   python tools/build_report.py --write --html befektetesi-alapok-riport.html
   python tools/build_report.py --write --html riport-b.html
   ```

További kapcsolók: `--src <mappa>` (alapért. a `Downloads`), `--dump out.json`.

## Validáció

A `--check` cellánként összeveti az újraépített adatot a HTML-be már beágyazottal.
A jelenlegi forrásokkal mind a 12 szekció **0 eltérést** ad (teljes paritás).

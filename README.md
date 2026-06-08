# 🏀 NBA Meccs Összefoglalók — Hub

Egy statikus, NBA-tematikájú gyűjtőoldal, ami a YouTube legnézettebb NBA
meccs-összefoglaló videóit jeleníti meg, **dátum szerint rendezve** (legfrissebb felül),
**beágyazva** az oldalra.

## Funkciók

- 📅 Dátum szerinti rendezés (legújabb / legrégebbi), valamint nézettség szerinti rendezés
- ▶️ Beágyazott YouTube-lejátszó hibatűréssel — ha egy videót a feltöltő nem enged
  beágyazni, az oldal automatikusan egy „Megnézés YouTube-on" gombra vált
- 🔄 Automatikus frissítés percenként (lejátszás közben nem szakítja meg a videót)
- 🔑 Opcionális élő betöltés a YouTube Data API v3-ról (a legnézettebb összefoglalók)

## Használat

Nyisd meg az [`index.html`](index.html) fájlt böngészőben — nincs build lépés,
nincs függőség. A teljes oldal egyetlen önálló HTML-fájl.

### Élő (legnézettebb) lista YouTube API-val

A jobb felső mezőbe írj be egy saját
[YouTube Data API v3](https://developers.google.com/youtube/v3/getting-started)
kulcsot, majd kattints az „Élő betöltés" gombra.

## Közzététel

Az oldal GitHub Pages-en van publikálva (lásd a repó *Settings → Pages* menüpontját).

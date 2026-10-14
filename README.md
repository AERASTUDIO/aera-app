# AERA Studio – app

Símaapp og tölvuapp AERA Studio fyrir verkefni, undirverk, fundi, tímaskráningu, skjöl og viðskiptavini.

- **Slóð:** https://aerastudio.github.io/aera-app
- **Hýsing:** GitHub Pages (útgáfa birtist sjálfkrafa við hverja breytingu á `main`)
- **Gögn:** Firebase Firestore, verkefnið `aera-app-cfdb3`, skjalið `app/state`
- **Innskráning:** Google, aðeins aerastudioice@gmail.com (sjá `firebase-config.js` og `firestore.rules`)

## Skrár
| Skrá | Hlutverk |
|---|---|
| `index.html` | Appið sjálft: útlit, skjámyndir og öll virkni |
| `firebase.js` | Innskráning og tenging við Firestore |
| `firebase-config.js` | Opinber auðkenni Firebase-verkefnisins og leyfð netföng |
| `firestore.rules` | Aðgangsreglur gagnagrunnsins (afrit; virku reglurnar eru í Firebase Console) |
| `sw.js` | Geymir appið á tækinu svo það virki án nets – hækka `VERSION` við breytingar |
| `manifest.webmanifest`, `icons/` | Nafn og tákn þegar appið er sett á heimaskjá |

Gögnin sjálf eru **aldrei** í þessari geymslu. Afrit fæst með „Sækja afrit af gögnum“ neðst í appinu.

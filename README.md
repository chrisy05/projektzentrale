# Projektzentrale

Deutschsprachige persönliche Projektübersicht mit einer gemeinsamen Aufgabenliste, Projektstatus, Wochenplanung, nachträglich erfasster Arbeitszeit und E.R.F.O.L.G.-Lebensplanung.

## Zwei Betriebsarten

- `dist/`: statische Vorschau. Speichert ausschließlich lokal im Browser. Kein gemeinsamer Sheets-Zugriff und keine automatischen Erinnerungen aus dem Webcode.
- `apps-script/`: dieselbe Oberfläche als private Google Apps Script Webapp. Google Sheets ist dann die gemeinsame Datenbasis. Noch nicht live bereitgestellt.

Die Vorschau kennzeichnet den lokalen Modus sichtbar. Persönliche Datensätze und Quelldateien gehören nicht in ein öffentliches Repository. `draft-data.json` gehört ausschließlich zur privaten Vorschau und wird nicht nach GitHub übertragen.

## Google Sheets aktivieren

1. Gewünschtes eigenes Google Sheet öffnen, **Erweiterungen → Apps Script**.
2. `apps-script/Code.gs`, `Index.html` und `appsscript.json` in das gebundene Scriptprojekt übernehmen.
3. Im Editor `setup_` einmal ausführen und Google-Zugriff bestätigen. Es legt `App_*`-Tabellen an. Bestehende Tabellen bleiben unverändert.
4. Bereitstellen → Neue Bereitstellung → Web-App. **Ausführen als: Nutzer, der auf die Web-App zugreift. Zugriff: Nur ich.** Nicht öffentlich bereitstellen.
5. Die neue Webapp öffnen. Der Banner muss **Google Sheets verbunden** zeigen. Einen eigenen Testschritt anlegen, im `App_Aufgaben`-Tab prüfen, abhaken und erneut laden.
6. Ab dann sind die `App_*`-Tabellen führend für Appdaten. Vorhandene alte Aufgaben/Projekte einmal kontrolliert übernehmen; nicht parallel zwei aktive Listen führen. Die UI-Sicherung lässt sich mit JSON importieren.

Die App prüft die eingeloggte Eigentümeradresse, verweigert unbekannte Benutzer und schützt gegen konkurrierende Änderungen über Prüfsummen. Vor Schreibvorgängen werden Zeilen in `App_Backups` gesichert; eine fehlgeschlagene Änderung wird zurückgesetzt. Keine Tokens oder Passwörter im Frontend. Kein offener API-Proxy.

## Entwicklung

`node --check dist/app.js`

`python3 scripts/build-apps-script.py`

Bei Änderungen am Frontend `Index.html` neu erzeugen. Apps Script muss erneut bereitgestellt werden. Die statische Vorschau kann separat zugangsgeschützt gehostet werden. GitHub Pages auf einem öffentlichen Repo schützt persönliche Inhalte nicht.

## Planung

Bis zu fünf Meilensteine pro Jahr über fünf Jahre; pro Meilenstein bis zu fünf Schritte empfohlen. Weitere operative Aufgaben werden in derselben Aufgabenliste erfasst. Kein Zwang, alle 125 Plätze zu füllen. Wochenbudget inklusive separat geführtem Schlaf und Puffer maximal 168 Stunden. Lebensblume als persönliches Sinnbild, keine Gesundheitsdiagnostik.

Arbeitszeiten manuell nachtragen: bestätigt oder geschätzt. Keine Start-/Stopp-Uhr. Erinnerungen benötigen eine separat eingerichtete Automation; der Browser läuft nicht automatisch im Hintergrund weiter.

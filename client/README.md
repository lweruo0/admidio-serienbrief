# Serienbrief-Client (Rust)

Windows-Programm mit grafischer Oberfläche, das aus einer Word-Serienbriefvorlage und einer
CSV-Datenquelle pro Datensatz eine PDF-Datei erzeugt. Es ist der Nachbau von
`com_mail_merge_class.py` und steuert Word wie das Python-Skript über COM-Automation.

## Voraussetzungen

- Windows mit installiertem Microsoft Word (COM-ProgID `Word.Application`).
- Rust-Toolchain (stable, Ziel `x86_64-pc-windows-msvc`).

## Bauen und starten

```powershell
cd client
cargo run --release
```

Die fertige Datei liegt danach unter `target\release\serienbrief-client.exe` und kann
ohne weitere Abhängigkeiten kopiert werden.

## Bedienung

1. **Serienbrief (Word)**: Vorlage mit den Seriendruckfeldern auswählen.
2. **Datenquelle (CSV)**: Export aus Admidio auswählen. Trennzeichen (`;`, `,` oder Tabulator)
   und Kodierung (UTF-8 oder ISO-8859-1) werden automatisch erkannt. Nach der Auswahl zeigt das
   Programm Datensatzanzahl, Spaltennamen und eine Vorschau des ersten Dateinamens.
3. **Zielordner**: wird beim Auswählen der Vorlage mit `<Ordner der Vorlage>\<Datum>-Rundschreiben`
   vorbelegt und bei Bedarf angelegt.
4. **Dateiname**: Muster mit Platzhaltern in geschweiften Klammern, Standard
   `rundschreiben-{Mitgliedsnummer}-{username}.pdf`. Jeder Platzhalter muss eine Spalte der
   Datenquelle sein; unbekannte Platzhalter werden rot angezeigt. In Dateinamen unzulässige
   Zeichen (`\ / : * ? " < > |`) werden durch `_` ersetzt.
5. **Parallele Word-Instanzen**: Anzahl der gleichzeitig laufenden Word-Prozesse (Standard 8).
   Die Datensätze werden in gleich große Blöcke aufgeteilt, jeder Block läuft in einer eigenen
   Word-Instanz mit eigener Kopie von Vorlage und Datenquelle im Temp-Ordner.
6. **Word-Prozesse beenden**: optional, entspricht dem `taskkill` im Python-Skript. Achtung:
   beendet auch offene Word-Fenster mit ungesicherten Änderungen.

### Datenquelle aus Admidio aktualisieren

Der aufklappbare Bereich **Datenquelle aus Admidio aktualisieren** lädt die CSV-Exporte direkt
aus Admidio:

- **Benutzername** und **Passwort** des Admidio-Kontos. Der Benutzername wird gespeichert, das
  Passwort nicht.
- **Export-URLs**, eine je Zeile. Das ist jeweils die URL des CSV-Exports einer Rollenliste,
  z. B. `https://<host>/modules/groups-roles/lists_show.php?list_uuid=...&role_list=...&mode=csv`.
  Zeilen, die mit `#` beginnen, werden ignoriert.
- **Spaltenzuordnung**: der Knopf **Bearbeiten …** öffnet ein Fenster mit einer Tabelle, in der
  links der Spaltenname steht, den der Serienbrief (die CSV) erwartet, und rechts der Name des
  Felds im Admidio-Export (z. B. `Mitgliedsnummer` ← `Mitgliedsnr.`, `username` ← `Benutzername`).
  Beim Download werden die Spaltenköpfe entsprechend umbenannt; nicht aufgeführte Spalten
  behalten ihren Namen aus Admidio. **Spalten der Datenquelle übernehmen** legt für jede Spalte
  der aktuell gewählten CSV eine Zeile an, in der nur noch der Admidio-Name anzupassen ist.
  Leere Zeilen werden ignoriert; halb ausgefüllte Zeilen und doppelte Namen werden rot gemeldet
  und verhindern den Download. Die Zuordnung wird mit den übrigen Einstellungen gespeichert.
- **Einmalpasswörter (WordPress)**, optional: Adresse der WordPress-Seite (z. B.
  `https://bfv-ehingen.de/`) und das **OTP-Passwort** des Endpunkts. Bleibt die Adresse leer,
  werden keine OTP-Spalten ergänzt. Das Häkchen **Einmalpasswörter auf dem Server neu erzeugen**
  setzt `clearotp=1` wie in `one_time_passwords_get.py`. Adresse, Passwort und Häkchen werden
  mit den Einstellungen gespeichert.

Ablauf beim Klick auf **CSV aus Admidio laden**: Anmeldeseite (`system/login.php`, aus der
ersten URL abgeleitet) abrufen, CSRF-Token und Session-Cookie übernehmen, anmelden, alle URLs
laden. Die Spaltenköpfe jedes Exports werden gemäß Spaltenzuordnung umbenannt, die Exporte
dann nach Spaltennamen vereinigt, vollständig identische Datensätze (Mitglied in mehreren
Rollen) nur einmal übernommen. Das Protokoll nennt die angewendeten Umbenennungen und warnt,
wenn ein zugeordnetes Admidio-Feld in keinem Export vorkam. Das Ergebnis wird als UTF-8 mit BOM
in die oben eingetragene Datenquelle geschrieben; ist dort noch nichts eingetragen, fragt das
Programm nach einer Zieldatei. Liefert eine URL kein CSV (fehlende Berechtigung, falsche URL),
bricht der Download mit Fehlermeldung ab.

Ist eine WordPress-Adresse eingetragen, ruft das Programm nach dem Admidio-Download die
OTP-Liste ab: `https://<wordpress-seite>/?getOTPlist=1&passwd=<OTP-Passwort>&nonce=<MD5 des
Datums JJJJMMTT>&clearotp=<0|1>`. Die Antwort ist ein JSON-Objekt mit dem Anmeldenamen als
Schlüssel und `[Anmeldename, Mitgliedsnummer, otp1, otp2]` als Wert. Jeder Datensatz erhält
die Spalten `otp1` und `otp2`; zugeordnet wird über die Spalte `username` (Anmeldename),
ersatzweise über `Mitgliedsnummer`. Datensätze ohne Treffer bekommen leere OTP-Spalten, das
Protokoll nennt ihre Anzahl. Ist die Antwort kein JSON (z. B. „Anfrage gesperrt“), bricht der
Download mit Fehlermeldung ab, die CSV wird dann nicht geschrieben.

Der Knopf **PDFs erzeugen** startet den Lauf. Fortschritt, vergangene und geschätzte Restzeit
sowie ein Protokoll der erzeugten Dateien werden angezeigt. Die zuletzt verwendeten Eingaben
werden beim Beenden gespeichert.

## Aufbau

| Datei              | Inhalt                                                                  |
|--------------------|-------------------------------------------------------------------------|
| `src/main.rs`      | Oberfläche (egui/eframe), Eingabeprüfung, Fortschrittsanzeige           |
| `src/merge.rs`     | CSV-Analyse, Platzhalter, Aufteilung auf Threads, Steuerung von Word     |
| `src/admidio.rs`   | Anmeldung bei Admidio, Download, Spaltenzuordnung, Zusammenführen der CSV-Exporte, OTP-Abruf von WordPress |
| `test/`            | Beispielvorlage `example.docx` und `example.csv` mit 800 Datensätzen     |
| `src/word.rs`      | Late-Binding-Wrapper um `IDispatch` (Property-Zugriff, Methodenaufrufe) |

## Tests

```powershell
cargo test
```

Ein kompletter Durchlauf mit echtem Word ist als ignorierter Test enthalten:

```powershell
$env:SERIENBRIEF_TEMPLATE = "C:\pfad\vorlage.docx"
$env:SERIENBRIEF_DATA = "C:\pfad\daten.csv"
$env:SERIENBRIEF_OUT = "C:\pfad\ausgabe"
cargo test full_merge_with_word -- --ignored --nocapture
```

Die CSV braucht dafür die Spalten `Mitgliedsnummer` und `username`.

Die Anmeldung gegen eine echte Admidio-Installation (Ablehnung eines falschen Passworts) prüft:

```powershell
$env:ADMIDIO_LIST_URL = "https://<host>/modules/groups-roles/lists_show.php?list_uuid=...&mode=csv"
cargo test admidio_login -- --ignored --nocapture
```

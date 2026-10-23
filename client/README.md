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

Pflichtfelder, die noch leer oder ungültig sind, sind leicht rosa hinterlegt (im dunklen
Design gedeckt rot). Sobald alle Felder einer Aktion ausgefüllt sind, wird der zugehörige
Knopf grün.

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

Unter den Pfadangaben stehen zwei Gruppen: **Admidio** (Anmeldung, Export-URL, Rolle,
Zeitraum, Spaltenzuordnung) und **WordPress (Einmalpasswörter)**. Die Gruppe Admidio lädt die
CSV-Exporte direkt aus Admidio:

- **Benutzername** und **Passwort** des Admidio-Kontos. Der Benutzername wird gespeichert, das
  Passwort nicht.
- **Export-URL**: die URL des CSV-Exports einer Rollenliste,
  z. B. `https://<host>/modules/groups-roles/lists_show.php?list_uuid=...&role_list=...&date_from=...&date_to=...&mode=csv`.
  Mehrere URLs sind weiterhin möglich, eine je Zeile; Zeilen, die mit `#` beginnen, werden
  ignoriert.
- **Rolle wählen** (Rollen-Knöpfe): der Knopf **Bearbeiten …** öffnet eine Tabelle mit
  Beschriftung (Rollenname) und dem Wert für den URL-Parameter `role_list`. In Admidio 5 ist
  das die UUID der Rolle; mehrere Rollen werden durch Komma getrennt. Für jede Zeile erscheint
  ein Knopf neben **Rolle wählen**. Ein Klick ersetzt in der Export-URL (bei mehreren Zeilen in
  jeder) den Parameter `role_list` durch den Wert des Knopfs; fehlt er, wird er ergänzt. Die
  übrigen Parameter `list_uuid`, `date_from`, `date_to` und `mode` bleiben unverändert. Der
  Download selbst wird danach wie gewohnt mit **CSV aus Admidio laden** gestartet. Der Knopf,
  dessen Wert gerade in der Export-URL steht, ist grün hervorgehoben. Leere Zeilen werden
  ignoriert; halb ausgefüllte Zeilen und doppelte Beschriftungen werden rot gemeldet. Die
  Knöpfe werden mit den übrigen Einstellungen gespeichert.
- **Zeitraum**: zeigt `date_from` und `date_to` aus der Export-URL. Der Knopf **Heute** setzt
  beide Parameter auf das aktuelle Datum (`JJJJ-MM-TT`) und ist grün, wenn die URL bereits auf
  heute zeigt. Fehlen die Parameter, nimmt Admidio selbst das heutige Datum.
- **Spaltenzuordnung**: der Knopf **Bearbeiten …** öffnet ein Fenster mit einer Tabelle, in der
  links der Spaltenname steht, den der Serienbrief (die CSV) erwartet, und rechts der Name des
  Felds im Admidio-Export (z. B. `Mitgliedsnummer` ← `Mitgliedsnr.`, `username` ← `Benutzername`).
  Beim Download werden die Spaltenköpfe entsprechend umbenannt; nicht aufgeführte Spalten
  behalten ihren Namen aus Admidio. **Spalten der Datenquelle übernehmen** legt für jede Spalte
  der aktuell gewählten CSV eine Zeile an, in der nur noch der Admidio-Name anzupassen ist.
  Leere Zeilen werden ignoriert; halb ausgefüllte Zeilen und doppelte Namen werden rot gemeldet
  und verhindern den Download. Die Zuordnung wird mit den übrigen Einstellungen gespeichert.
Die Gruppe **WordPress (Einmalpasswörter)** ist optional:

- Das Häkchen **Einmalpasswörter beim Download abrufen** schaltet den Abruf ein. Es wird
  bewusst nicht gespeichert, damit die OTPs nicht versehentlich bei jedem Download geholt
  werden. Solange es aus ist, bleiben die Felder darunter gesperrt und die CSV bekommt keine
  OTP-Spalten.
- **WordPress-Adresse** (z. B. `https://bfv-ehingen.de/`) und das **OTP-Passwort** des
  Endpunkts. Beide werden mit den Einstellungen gespeichert.
- Ist das Häkchen gesetzt, fragt das Programm beim Klick auf **CSV aus Admidio laden** in
  einem Dialog, ob die bestehenden Einmalpasswörter auf dem Server gelöscht und neu erzeugt
  werden sollen: **Ja** setzt `clearotp=1` wie in `one_time_passwords_get.py`, die alten OTPs
  werden ungültig. **Nein** (`clearotp=0`) lässt die alten OTPs gültig und erzeugt zusätzlich
  neue. **Abbrechen** startet keinen Download. Vorbelegt ist **Nein**, Enter löscht also nie
  versehentlich die alten OTPs.

Ablauf beim Klick auf **CSV aus Admidio laden**: Anmeldeseite (`system/login.php`, aus der
ersten URL abgeleitet) abrufen, CSRF-Token und Session-Cookie übernehmen, anmelden, alle URLs
laden. Die Spaltenköpfe jedes Exports werden gemäß Spaltenzuordnung umbenannt, die Exporte
dann nach Spaltennamen vereinigt, vollständig identische Datensätze (Mitglied in mehreren
Rollen) nur einmal übernommen. Das Protokoll nennt die angewendeten Umbenennungen und warnt,
wenn ein zugeordnetes Admidio-Feld in keinem Export vorkam. Das Ergebnis wird als UTF-8 ohne BOM
mit Semikolon als Spaltentrennzeichen in die oben eingetragene CSV-Datenquelle geschrieben;
zusätzlich erzeugt das Programm im selben Ordner eine gleichnamige Excel-Datei mit der Endung
`.xlsx`. Ist noch keine Zieldatei eingetragen, fragt das Programm nach einem CSV-Zielpfad.
Liefert eine URL kein CSV (fehlende Berechtigung, falsche URL), bricht der Download mit
Fehlermeldung ab.

Ist der Abruf der Einmalpasswörter eingeschaltet, ruft das Programm nach dem Admidio-Download
die OTP-Liste ab: `https://<wordpress-seite>/?getOTPlist=1&passwd=<OTP-Passwort>&nonce=<MD5 des
Datums JJJJMMTT>&clearotp=<0|1>`. Die Antwort ist ein JSON-Objekt mit dem Anmeldenamen als
Schlüssel und `[Anmeldename, Mitgliedsnummer, otp1, otp2]` als Wert. Jeder Datensatz erhält
die Spalten `otp1` und `otp2`; zugeordnet wird über die Spalte `username` (Anmeldename),
ersatzweise über `Mitgliedsnummer`. Datensätze ohne Treffer bekommen leere OTP-Spalten, das
Protokoll nennt ihre Anzahl. Ist die Antwort kein JSON (z. B. „Anfrage gesperrt“), bricht der
Download mit Fehlermeldung ab, die CSV wird dann nicht geschrieben.

Der Knopf **Word öffnen** daneben startet Word sichtbar, öffnet die gewählte
Serienbrief-Vorlage und hängt die eingetragene Datenquelle als Seriendruck-Datenquelle an,
genau wie beim Lauf. So lassen sich Seriendruckfelder und Vorschau direkt in Word prüfen.
Fehlt die Datenquelle oder ist sie unlesbar, wird nur die Vorlage geöffnet; das Protokoll
sagt es. Word bleibt danach geöffnet, Änderungen an der Vorlage speichert man in Word
selbst. Der Knopf ist nur aktiv, solange die Vorlage vorhanden ist und kein Lauf läuft.
**Zielordner öffnen** zeigt den Zielordner im Explorer; der Knopf ist aktiv, sobald der Ordner
existiert, also auch während eines Laufs, um die entstehenden PDFs zu sehen.

Der Knopf **PDFs erzeugen** startet den Lauf. Fortschritt, vergangene und geschätzte Restzeit
sowie ein Protokoll der erzeugten Dateien werden angezeigt. Die zuletzt verwendeten Eingaben
werden beim Beenden gespeichert.

## Aufbau

| Datei              | Inhalt                                                                  |
|--------------------|-------------------------------------------------------------------------|
| `src/main.rs`      | Oberfläche (egui/eframe), Eingabeprüfung, Fortschrittsanzeige           |
| `src/merge.rs`     | CSV-Analyse, Platzhalter, Aufteilung auf Threads, Steuerung von Word     |
| `src/admidio.rs`   | Anmeldung bei Admidio, Download, Spaltenzuordnung, Rollen-Knöpfe und Zeitraum (`role_list`, `date_from`, `date_to` in der URL ersetzen), Zusammenführen der CSV-Exporte, OTP-Abruf von WordPress |
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

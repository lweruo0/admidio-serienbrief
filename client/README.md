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

Der Knopf **PDFs erzeugen** startet den Lauf. Fortschritt, vergangene und geschätzte Restzeit
sowie ein Protokoll der erzeugten Dateien werden angezeigt. Die zuletzt verwendeten Eingaben
werden beim Beenden gespeichert.

## Aufbau

| Datei              | Inhalt                                                                  |
|--------------------|-------------------------------------------------------------------------|
| `src/main.rs`      | Oberfläche (egui/eframe), Eingabeprüfung, Fortschrittsanzeige           |
| `src/merge.rs`     | CSV-Analyse, Platzhalter, Aufteilung auf Threads, Steuerung von Word     |
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

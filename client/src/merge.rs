//! Serienbrief-Logik: Datenquelle analysieren, Datensätze auf Threads verteilen und jeden
//! Datensatz über Word als einzelne PDF-Datei ausgeben.

use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::mpsc::Sender;
use std::sync::Arc;
use std::thread;
use std::time::{Duration, Instant};

use anyhow::{bail, Context, Result};
use windows::core::VARIANT;

use crate::word::{ComApartment, ComObject};

/// Word-Konstante wdSendToNewDocument.
const WD_SEND_TO_NEW_DOCUMENT: i32 = 0;
/// Word-Konstante wdFormatPDF.
const WD_FORMAT_PDF: i32 = 17;
/// Word-Konstante wdDoNotSaveChanges.
const WD_DO_NOT_SAVE_CHANGES: i32 = 0;
/// Word-Konstante wdAlertsNone.
const WD_ALERTS_NONE: i32 = 0;

/// Alle Einstellungen für einen Serienbrief-Lauf.
#[derive(Clone, Debug)]
pub struct MergeConfig {
    /// Word-Vorlage mit den Seriendruckfeldern.
    pub template: PathBuf,
    /// CSV-Datenquelle (ein Datensatz pro Zeile, erste Zeile Spaltennamen).
    pub data_source: PathBuf,
    /// Zielordner für die PDF-Dateien.
    pub output_dir: PathBuf,
    /// Dateinamen-Muster mit Platzhaltern in geschweiften Klammern,
    /// z. B. `rundschreiben-{Mitgliedsnummer}-{username}.pdf`.
    pub file_pattern: String,
    /// Anzahl paralleler Word-Instanzen.
    pub thread_count: usize,
    /// Vor dem Start alle laufenden Word-Prozesse beenden.
    pub kill_word: bool,
}

/// Fortschrittsmeldungen der Arbeits-Threads an die Oberfläche.
#[derive(Debug)]
pub enum Progress {
    /// Gesamtzahl der Datensätze steht fest.
    Total(usize),
    /// Eine Textzeile für das Protokoll.
    Log(String),
    /// Ein Datensatz wurde erfolgreich als PDF gespeichert.
    RecordDone,
    /// Ein Datensatz konnte nicht verarbeitet werden.
    RecordFailed,
    /// Der gesamte Lauf ist beendet.
    Finished(Duration),
}

/// Kennzahlen einer CSV-Datenquelle für die Vorschau in der Oberfläche.
#[derive(Clone, Debug, Default)]
pub struct DataSourceInfo {
    /// Spaltennamen aus der Kopfzeile.
    pub fields: Vec<String>,
    /// Anzahl der Datensätze ohne Kopfzeile.
    pub record_count: usize,
    /// Werte des ersten Datensatzes, zur Vorschau des Dateinamens.
    pub first_record: HashMap<String, String>,
}

/// Liest Kopfzeile, Datensatzanzahl und ersten Datensatz einer CSV-Datei.
pub fn inspect_data_source(path: &Path) -> Result<DataSourceInfo> {
    let bytes = fs::read(path)
        .with_context(|| format!("Datenquelle '{}' kann nicht gelesen werden", path.display()))?;
    let bytes = bytes.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(&bytes);

    let mut reader = csv::ReaderBuilder::new()
        .delimiter(detect_delimiter(bytes))
        .flexible(true)
        .has_headers(true)
        .from_reader(bytes);

    let fields: Vec<String> = reader
        .byte_headers()
        .context("Kopfzeile der Datenquelle ist ungültig")?
        .iter()
        .map(|field| decode_text(field).trim().to_string())
        .collect();

    let mut info = DataSourceInfo { fields, ..Default::default() };
    for record in reader.byte_records() {
        let record = record.context("Datensatz in der Datenquelle ist ungültig")?;
        if record.iter().all(|field| field.trim_ascii().is_empty()) {
            continue;
        }
        if info.record_count == 0 {
            for (name, value) in info.fields.iter().zip(record.iter()) {
                info.first_record.insert(name.clone(), decode_text(value));
            }
        }
        info.record_count += 1;
    }
    Ok(info)
}

/// Bestimmt das Trennzeichen anhand der Kopfzeile (Semikolon, Komma oder Tabulator).
fn detect_delimiter(bytes: &[u8]) -> u8 {
    let first_line = bytes.split(|byte| *byte == b'\n').next().unwrap_or(&[]);
    // Bei Gleichstand gewinnt der letzte Kandidat, daher steht das Semikolon am Ende.
    [b'\t', b',', b';']
        .into_iter()
        .max_by_key(|candidate| first_line.iter().filter(|byte| *byte == candidate).count())
        .unwrap_or(b';')
}

/// Dekodiert CSV-Inhalt als UTF-8, mit Rückfall auf ISO-8859-1.
fn decode_text(bytes: &[u8]) -> String {
    match std::str::from_utf8(bytes) {
        Ok(text) => text.to_string(),
        Err(_) => bytes.iter().map(|&byte| byte as char).collect(),
    }
}

/// Liefert alle Platzhalternamen eines Dateinamen-Musters in Reihenfolge ihres Auftretens.
pub fn placeholders(pattern: &str) -> Vec<String> {
    let mut names: Vec<String> = Vec::new();
    let mut rest = pattern;
    while let Some(start) = rest.find('{') {
        let after = &rest[start + 1..];
        let Some(end) = after.find('}') else {
            break;
        };
        let name = after[..end].trim();
        if !name.is_empty() && !names.iter().any(|known| known == name) {
            names.push(name.to_string());
        }
        rest = &after[end + 1..];
    }
    names
}

/// Ersetzt die Platzhalter eines Musters durch die übergebenen Werte.
pub fn render(pattern: &str, values: &HashMap<String, String>) -> String {
    let mut result = pattern.to_string();
    for name in placeholders(pattern) {
        if let Some(value) = values.get(&name) {
            result = result.replace(&format!("{{{name}}}"), &sanitize_filename_part(value));
        }
    }
    result
}

/// Entfernt Zeichen, die in Windows-Dateinamen nicht erlaubt sind.
fn sanitize_filename_part(value: &str) -> String {
    value
        .trim()
        .chars()
        .map(|c| match c {
            '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|' => '_',
            c if c.is_control() => '_',
            c => c,
        })
        .collect()
}

/// Formatiert eine Dauer als `MM:SS`.
pub fn format_duration(duration: Duration) -> String {
    let seconds = duration.as_secs();
    format!("{:02}:{:02}", seconds / 60, seconds % 60)
}

/// Startet den Serienbrief-Lauf in einem Hintergrund-Thread. Der Fortschritt wird über den
/// übergebenen Kanal gemeldet; der Aufruf kehrt sofort zurück.
pub fn start(config: MergeConfig, tx: Sender<Progress>) {
    thread::spawn(move || {
        let started = Instant::now();
        if let Err(error) = run(config, &tx) {
            let _ = tx.send(Progress::Log(format!("Abbruch: {error:#}")));
        }
        let _ = tx.send(Progress::Finished(started.elapsed()));
    });
}

/// Führt den Lauf aus: Vorbereitung, Verteilung auf Threads, Warten auf das Ende.
fn run(config: MergeConfig, tx: &Sender<Progress>) -> Result<()> {
    if config.kill_word {
        kill_word_processes();
        let _ = tx.send(Progress::Log("Laufende Word-Prozesse beendet.".to_string()));
    }

    let info = inspect_data_source(&config.data_source)?;
    if info.record_count == 0 {
        bail!("Die Datenquelle enthält keine Datensätze");
    }
    let unknown: Vec<String> = placeholders(&config.file_pattern)
        .into_iter()
        .filter(|name| !info.fields.contains(name))
        .collect();
    if !unknown.is_empty() {
        bail!("Unbekannte Platzhalter im Dateinamen: {}", unknown.join(", "));
    }

    fs::create_dir_all(&config.output_dir).with_context(|| {
        format!("Zielordner '{}' kann nicht angelegt werden", config.output_dir.display())
    })?;

    let _ = tx.send(Progress::Total(info.record_count));
    let _ = tx.send(Progress::Log(format!(
        "{} Datensätze, Ausgabe nach {}",
        info.record_count,
        config.output_dir.display()
    )));

    // Datensatznummern (1-basiert, wie in Word) in möglichst gleich große Blöcke aufteilen.
    let thread_count = config.thread_count.clamp(1, info.record_count);
    let indices: Vec<usize> = (1..=info.record_count).collect();
    let chunk_size = indices.len().div_ceil(thread_count);
    let config = Arc::new(config);

    let workers: Vec<_> = indices
        .chunks(chunk_size)
        .enumerate()
        .map(|(index, chunk)| {
            let config = Arc::clone(&config);
            let chunk = chunk.to_vec();
            let tx = tx.clone();
            thread::spawn(move || {
                if let Err(error) = merge_chunk(index, &config, &chunk, &tx) {
                    let _ = tx.send(Progress::Log(format!("Thread {index}: {error:#}")));
                }
            })
        })
        .collect();

    for worker in workers {
        let _ = worker.join();
    }
    Ok(())
}

/// Beendet alle WINWORD.EXE-Prozesse, damit keine alten Instanzen den Seriendruck stören.
fn kill_word_processes() {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x0800_0000;
    let _ = std::process::Command::new("taskkill")
        .args(["/f", "/im", "WINWORD.EXE"])
        .creation_flags(CREATE_NO_WINDOW)
        .output();
}

/// Verarbeitet einen Block von Datensätzen in einer eigenen Word-Instanz.
///
/// Vorlage und Datenquelle werden vorher in den Temp-Ordner kopiert, damit sich die parallel
/// laufenden Word-Instanzen nicht gegenseitig die Dateien sperren.
fn merge_chunk(index: usize, config: &MergeConfig, chunk: &[usize], tx: &Sender<Progress>) -> Result<()> {
    let temp_dir = std::env::temp_dir();
    let extension = config
        .template
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("docx");
    let template_copy = temp_dir.join(format!("rundschreiben_{index}.{extension}"));
    let data_copy = temp_dir.join(format!("data_{index}.csv"));

    fs::copy(&config.template, &template_copy).context("Vorlage kann nicht kopiert werden")?;
    fs::copy(&config.data_source, &data_copy).context("Datenquelle kann nicht kopiert werden")?;

    let outcome = merge_with_word(index, config, chunk, &template_copy, &data_copy, tx);

    let _ = fs::remove_file(&template_copy);
    let _ = fs::remove_file(&data_copy);
    outcome
}

/// Steuert Word: Vorlage öffnen, Datenquelle anhängen und jeden Datensatz als PDF speichern.
fn merge_with_word(
    index: usize,
    config: &MergeConfig,
    chunk: &[usize],
    template_copy: &Path,
    data_copy: &Path,
    tx: &Sender<Progress>,
) -> Result<()> {
    let _apartment = ComApartment::new()?;
    let app = ComObject::create("Word.Application")?;
    let _ = app.put("Visible", false);
    let _ = app.put("DisplayAlerts", WD_ALERTS_NONE);

    let outcome = (|| -> Result<()> {
        let documents = app.get_object("Documents")?;
        let template = documents.call_object("Open", vec![path_variant(template_copy)])?;
        let mail_merge = template.get_object("MailMerge")?;
        mail_merge.call("OpenDataSource", vec![path_variant(data_copy)])?;
        mail_merge.put("Destination", WD_SEND_TO_NEW_DOCUMENT)?;
        let data_source = mail_merge.get_object("DataSource")?;
        let names = placeholders(&config.file_pattern);

        for &record in chunk {
            match merge_record(&app, &mail_merge, &data_source, &names, record, config) {
                Ok(target) => {
                    let _ = tx.send(Progress::Log(format!("[{index}] {}", target.display())));
                    let _ = tx.send(Progress::RecordDone);
                }
                Err(error) => {
                    let _ = tx.send(Progress::Log(format!("[{index}] Datensatz {record}: {error:#}")));
                    let _ = tx.send(Progress::RecordFailed);
                }
            }
        }

        template.call("Close", vec![VARIANT::from(WD_DO_NOT_SAVE_CHANGES)])?;
        Ok(())
    })();

    let _ = app.call("Quit", vec![VARIANT::from(WD_DO_NOT_SAVE_CHANGES)]);
    outcome
}

/// Führt den Seriendruck für genau einen Datensatz aus und speichert das Ergebnis als PDF.
fn merge_record(
    app: &ComObject,
    mail_merge: &ComObject,
    data_source: &ComObject,
    names: &[String],
    record: usize,
    config: &MergeConfig,
) -> Result<PathBuf> {
    let record_number = record as i32;
    data_source.put("FirstRecord", record_number)?;
    data_source.put("LastRecord", record_number)?;
    data_source.put("ActiveRecord", record_number)?;

    let mut values = HashMap::new();
    for name in names {
        let field = data_source.call_object("DataFields", vec![VARIANT::from(name.as_str())])?;
        values.insert(name.clone(), field.get_string("Value")?);
    }
    let target = config.output_dir.join(render(&config.file_pattern, &values));

    mail_merge.call("Execute", Vec::new())?;
    let result_document = app.get_object("ActiveDocument")?;
    let saved = result_document.call(
        "SaveAs",
        vec![path_variant(&target), VARIANT::from(WD_FORMAT_PDF)],
    );
    let _ = result_document.call("Close", vec![VARIANT::from(WD_DO_NOT_SAVE_CHANGES)]);
    saved?;
    Ok(target)
}

/// Wandelt einen Pfad in einen VARIANT-String für Word.
fn path_variant(path: &Path) -> VARIANT {
    VARIANT::from(path.to_string_lossy().as_ref())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn placeholders_are_extracted_in_order() {
        let names = placeholders("rundschreiben-{Mitgliedsnummer}-{username}.pdf");
        assert_eq!(names, vec!["Mitgliedsnummer".to_string(), "username".to_string()]);
    }

    #[test]
    fn render_replaces_and_sanitizes() {
        let mut values = HashMap::new();
        values.insert("Mitgliedsnummer".to_string(), "4711".to_string());
        values.insert("username".to_string(), "max/muster:mann".to_string());
        let name = render("rundschreiben-{Mitgliedsnummer}-{username}.pdf", &values);
        assert_eq!(name, "rundschreiben-4711-max_muster_mann.pdf");
    }

    /// Kompletter Durchlauf mit echtem Word. Läuft nur, wenn die Umgebungsvariablen
    /// `SERIENBRIEF_TEMPLATE`, `SERIENBRIEF_DATA` und `SERIENBRIEF_OUT` gesetzt sind:
    /// `cargo test -- --ignored`.
    #[test]
    #[ignore]
    fn full_merge_with_word() {
        let config = MergeConfig {
            template: PathBuf::from(std::env::var("SERIENBRIEF_TEMPLATE").unwrap()),
            data_source: PathBuf::from(std::env::var("SERIENBRIEF_DATA").unwrap()),
            output_dir: PathBuf::from(std::env::var("SERIENBRIEF_OUT").unwrap()),
            file_pattern: "rundschreiben-{Mitgliedsnummer}-{username}.pdf".to_string(),
            thread_count: 2,
            kill_word: false,
        };
        let (tx, rx) = std::sync::mpsc::channel();
        start(config.clone(), tx);
        let mut done = 0;
        let mut failed = 0;
        for message in rx {
            match message {
                Progress::Log(line) => println!("{line}"),
                Progress::RecordDone => done += 1,
                Progress::RecordFailed => failed += 1,
                Progress::Finished(duration) => println!("Dauer {}", format_duration(duration)),
                Progress::Total(total) => println!("{total} Datensätze"),
            }
        }
        assert_eq!(failed, 0);
        assert_eq!(done, inspect_data_source(&config.data_source).unwrap().record_count);
        assert!(config.output_dir.join("rundschreiben-1001-max.muster.pdf").is_file());
        assert!(config.output_dir.join("rundschreiben-1003-hans_ole.pdf").is_file());
    }

    #[test]
    fn delimiter_detection_prefers_semicolon_on_tie() {
        assert_eq!(detect_delimiter(b"a;b;c\n1;2;3"), b';');
        assert_eq!(detect_delimiter(b"a,b,c\n1,2,3"), b',');
        assert_eq!(detect_delimiter(b"a\tb\n1\t2"), b'\t');
        assert_eq!(detect_delimiter(b"a\n1"), b';');
    }
}

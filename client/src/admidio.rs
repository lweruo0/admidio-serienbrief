//! Aktualisierung der Datenquelle aus Admidio: Anmeldung per Benutzername und Passwort,
//! Download der CSV-Exporte mehrerer Rollenlisten und Zusammenführen zu einer Datei.
//! Optional werden anschließend Einmalpasswörter (OTP) von der WordPress-Seite geholt und
//! als Spalten `otp1`/`otp2` ergänzt (Nachbau von `one_time_passwords_get.py`).

use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::PathBuf;
use std::sync::mpsc::Sender;
use std::thread;
use std::time::Duration;

use anyhow::{anyhow, bail, Context, Result};
use reqwest::blocking::Client;
use reqwest::header::{CONTENT_DISPOSITION, CONTENT_TYPE};
use rust_xlsxwriter::Workbook;

use crate::merge::{decode_text, detect_delimiter};

/// Einstellungen für den Download aus Admidio.
#[derive(Clone, Debug)]
pub struct AdmidioConfig {
    /// Anmeldename in Admidio.
    pub username: String,
    /// Passwort in Admidio.
    pub password: String,
    /// CSV-Export-URLs (`lists_show.php?...&mode=csv`), eine je Rolle.
    pub urls: Vec<String>,
    /// Zieldatei für die zusammengeführte CSV.
    pub target: PathBuf,
    /// Zuordnung der CSV-Spalten zu den Feldnamen im Admidio-Export.
    pub columns: Vec<ColumnMapping>,
    /// Abruf der Einmalpasswörter von WordPress; `None` lässt die OTP-Spalten weg.
    pub otp: Option<OtpConfig>,
}

/// Einstellungen für den Abruf der Einmalpasswörter (OTP) von der WordPress-Seite.
#[derive(Clone, Debug)]
pub struct OtpConfig {
    /// Adresse der WordPress-Seite, z. B. `https://bfv-ehingen.de/`.
    pub site_url: String,
    /// Passwort des OTP-Endpunkts (Parameter `passwd`).
    pub password: String,
    /// Bestehende Einmalpasswörter auf dem Server löschen und neu erzeugen (`clearotp=1`).
    pub clear: bool,
}

/// Name der CSV-Spalte mit dem Anmeldenamen; über sie werden die OTPs zugeordnet.
pub const USERNAME_COLUMN: &str = "username";
/// Name der CSV-Spalte mit der Mitgliedsnummer; zweiter Schlüssel für die OTP-Zuordnung.
pub const MEMBER_NUMBER_COLUMN: &str = "Mitgliedsnummer";
/// Namen der Spalten, in die die Einmalpasswörter geschrieben werden.
pub const OTP_COLUMNS: [&str; 2] = ["otp1", "otp2"];

/// Ein Eintrag der OTP-Liste von WordPress.
#[derive(Clone, Debug, PartialEq)]
pub struct OtpEntry {
    /// Anmeldename des Mitglieds.
    pub login: String,
    /// Mitgliedsnummer.
    pub member_number: String,
    /// Erstes Einmalpasswort.
    pub otp1: String,
    /// Zweites Einmalpasswort.
    pub otp2: String,
}

/// Baut die Abruf-URL der OTP-Liste. Die Nonce ist der MD5-Hash des Datums `JJJJMMTT`,
/// wie es der WordPress-Endpunkt erwartet.
pub fn otp_url(config: &OtpConfig, date: &str) -> Result<String> {
    let nonce = format!("{:x}", md5::compute(date.as_bytes()));
    let clear = if config.clear { "1" } else { "0" };
    let base = format!("{}/", config.site_url.trim().trim_end_matches('/'));
    let url = reqwest::Url::parse_with_params(
        &base,
        [
            ("getOTPlist", "1"),
            ("passwd", config.password.trim()),
            ("nonce", nonce.as_str()),
            ("clearotp", clear),
        ],
    )
    .with_context(|| format!("WordPress-Adresse '{}' ist keine gültige URL", config.site_url))?;
    Ok(url.to_string())
}

/// Wertet die JSON-Antwort des OTP-Endpunkts aus. Erwartet wird ein Objekt mit dem
/// Anmeldenamen als Schlüssel und `[Anmeldename, Mitgliedsnummer, otp1, otp2]` als Wert.
pub fn parse_otp_list(body: &str) -> Result<Vec<OtpEntry>> {
    let json: serde_json::Value = serde_json::from_str(body.trim()).map_err(|_| {
        let snippet: String = body.trim().chars().take(120).collect();
        anyhow!("Antwort des OTP-Endpunkts ist kein JSON (Anfrage gesperrt?): '{snippet}'")
    })?;
    let object = json
        .as_object()
        .ok_or_else(|| anyhow!("OTP-Liste ist kein JSON-Objekt"))?;
    let mut entries = Vec::with_capacity(object.len());
    for (key, value) in object {
        let fields: Vec<String> = value
            .as_array()
            .ok_or_else(|| anyhow!("OTP-Eintrag '{key}' ist keine Liste"))?
            .iter()
            .map(|field| match field {
                serde_json::Value::String(text) => text.trim().to_string(),
                other => other.to_string(),
            })
            .collect();
        if fields.len() < 4 {
            bail!("OTP-Eintrag '{key}' hat nur {} statt 4 Felder", fields.len());
        }
        entries.push(OtpEntry {
            login: if fields[0].is_empty() { key.clone() } else { fields[0].clone() },
            member_number: fields[1].clone(),
            otp1: fields[2].clone(),
            otp2: fields[3].clone(),
        });
    }
    Ok(entries)
}

/// Zuordnung einer Spalte der Serienbrief-CSV zu einem Feld des Admidio-Exports.
#[derive(Clone, Debug, Default, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(default)]
pub struct ColumnMapping {
    /// Spaltenname, wie ihn die Word-Vorlage bzw. die CSV erwartet.
    pub csv: String,
    /// Spaltenname im CSV-Export von Admidio.
    pub admidio: String,
}

/// Prüft die Spaltenzuordnung und liefert die Umbenennungen (Admidio-Feld → CSV-Spalte).
/// Zeilen, in denen beide Felder leer sind, werden übergangen; halb ausgefüllte Zeilen und
/// doppelte Namen sind Fehler.
pub fn column_renames(mappings: &[ColumnMapping]) -> Result<HashMap<String, String>> {
    let mut renames = HashMap::new();
    let mut csv_names = HashSet::new();
    for (index, mapping) in mappings.iter().enumerate() {
        let csv = mapping.csv.trim();
        let admidio = mapping.admidio.trim();
        if csv.is_empty() && admidio.is_empty() {
            continue;
        }
        if csv.is_empty() || admidio.is_empty() {
            bail!("Spaltenzuordnung, Zeile {}: beide Felder müssen ausgefüllt sein", index + 1);
        }
        if !csv_names.insert(csv.to_string()) {
            bail!("Spaltenzuordnung: CSV-Spalte '{csv}' ist mehrfach angegeben");
        }
        if renames.insert(admidio.to_string(), csv.to_string()).is_some() {
            bail!("Spaltenzuordnung: Admidio-Feld '{admidio}' ist mehrfach angegeben");
        }
    }
    Ok(renames)
}

/// Ein konfigurierbarer Knopf, der die Export-URLs mit einer bestimmten Rolle lädt.
#[derive(Clone, Debug, Default, PartialEq, serde::Serialize, serde::Deserialize)]
#[serde(default)]
pub struct RoleButton {
    /// Beschriftung des Knopfs, z. B. „Vorstand“ oder „Alle Mitglieder“.
    pub name: String,
    /// Wert des URL-Parameters `role_list`: in Admidio 5 eine oder mehrere Rollen-UUIDs,
    /// durch Komma getrennt.
    pub role_list: String,
}

/// Prüft die Rollen-Knöpfe und liefert die ausgefüllten Einträge in Reihenfolge.
/// Zeilen, in denen beide Felder leer sind, werden übergangen; halb ausgefüllte Zeilen und
/// doppelte Beschriftungen sind Fehler.
pub fn role_buttons(buttons: &[RoleButton]) -> Result<Vec<RoleButton>> {
    let mut result = Vec::new();
    let mut names = HashSet::new();
    for (index, button) in buttons.iter().enumerate() {
        let name = button.name.trim();
        let role_list = button.role_list.trim();
        if name.is_empty() && role_list.is_empty() {
            continue;
        }
        if name.is_empty() || role_list.is_empty() {
            bail!("Rollen-Knöpfe, Zeile {}: Beschriftung und role_list müssen ausgefüllt sein", index + 1);
        }
        if !names.insert(name.to_string()) {
            bail!("Rollen-Knöpfe: Beschriftung '{name}' ist mehrfach angegeben");
        }
        result.push(RoleButton {
            name: name.to_string(),
            role_list: role_list.to_string(),
        });
    }
    Ok(result)
}

/// Ersetzt (oder ergänzt) die angegebenen Query-Parameter einer Listen-URL; alle anderen
/// Parameter wie `list_uuid` und `mode` bleiben in ihrer Reihenfolge erhalten.
pub fn with_query_params(list_url: &str, params: &[(&str, &str)]) -> Result<String> {
    let mut url = reqwest::Url::parse(list_url.trim())
        .with_context(|| format!("'{list_url}' ist keine gültige URL"))?;
    let pairs: Vec<(String, String)> = url
        .query_pairs()
        .filter(|(key, _)| !params.iter().any(|(name, _)| key == name))
        .map(|(key, value)| (key.into_owned(), value.into_owned()))
        .collect();
    {
        let mut query = url.query_pairs_mut();
        query.clear();
        for (key, value) in &pairs {
            query.append_pair(key, value);
        }
        for (name, value) in params {
            query.append_pair(name, value.trim());
        }
    }
    Ok(url.to_string())
}

/// Ersetzt die angegebenen Query-Parameter in jeder URL-Zeile des Textfelds; Leerzeilen und
/// Kommentarzeilen (`#`) bleiben unverändert erhalten.
pub fn replace_query_params(text: &str, params: &[(&str, &str)]) -> Result<String> {
    let mut lines = Vec::new();
    for line in text.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            lines.push(line.to_string());
        } else {
            lines.push(with_query_params(trimmed, params)?);
        }
    }
    Ok(lines.join("\n"))
}

/// Ersetzt `role_list` in jeder URL-Zeile des Textfelds.
pub fn replace_role_list(text: &str, role_list: &str) -> Result<String> {
    replace_query_params(text, &[("role_list", role_list)])
}

/// Liest einen Query-Parameter aus der ersten URL-Zeile des Textfelds. Liefert `None`,
/// wenn keine URL eingetragen, die URL ungültig oder der Parameter nicht vorhanden ist.
pub fn query_param(text: &str, name: &str) -> Option<String> {
    let first = parse_urls(text).into_iter().next()?;
    let url = reqwest::Url::parse(&first).ok()?;
    url.query_pairs()
        .find(|(key, _)| key == name)
        .map(|(_, value)| value.trim().to_string())
}

/// Meldungen des Download-Threads an die Oberfläche.
#[derive(Debug)]
pub enum DownloadEvent {
    /// Eine Textzeile für das Protokoll.
    Log(String),
    /// Download beendet, mit Pfad der geschriebenen Datei oder Fehlermeldung.
    Finished(Result<PathBuf, String>),
}

/// Startet den Download in einem Hintergrund-Thread; der Aufruf kehrt sofort zurück.
pub fn start(config: AdmidioConfig, tx: Sender<DownloadEvent>) {
    thread::spawn(move || {
        let result = download(&config, &tx).map_err(|error| format!("{error:#}"));
        let _ = tx.send(DownloadEvent::Finished(result));
    });
}

/// Zerlegt den Inhalt des URL-Textfelds (eine URL je Zeile) in eine Liste.
pub fn parse_urls(text: &str) -> Vec<String> {
    text.lines()
        .map(str::trim)
        .filter(|line| !line.is_empty() && !line.starts_with('#'))
        .map(str::to_string)
        .collect()
}

/// Leitet aus einer Listen-URL die URL der Anmeldeseite ab
/// (`https://host/pfad/modules/...` wird zu `https://host/pfad/system/login.php`).
pub fn login_url(list_url: &str) -> Result<String> {
    let position = list_url
        .find("/modules/")
        .ok_or_else(|| anyhow!("URL '{list_url}' enthält kein '/modules/' und ist keine Admidio-Listen-URL"))?;
    Ok(format!("{}/system/login.php", &list_url[..position]))
}

/// Liest den CSRF-Token aus dem HTML der Anmeldeseite.
pub fn csrf_token(html: &str) -> Option<String> {
    let start = html.find("adm_csrf_token")?;
    let rest = &html[start..];
    let end_of_tag = rest.find('>')?;
    let tag = &rest[..end_of_tag];
    let value_start = tag.find("value=\"")? + "value=\"".len();
    let value_end = tag[value_start..].find('"')?;
    let token = &tag[value_start..value_start + value_end];
    (!token.is_empty()).then(|| token.to_string())
}

/// Kürzt eine Admidio-Fehlermeldung um Dateipfad, Zeilennummer und Stacktrace.
fn clean_message(message: &str) -> String {
    let mut text = message;
    for marker in [" in /", "<br", "\n"] {
        if let Some(position) = text.find(marker) {
            text = &text[..position];
        }
    }
    text.trim().to_string()
}

fn log(tx: &Sender<DownloadEvent>, line: impl Into<String>) {
    let _ = tx.send(DownloadEvent::Log(line.into()));
}

/// Meldet sich an, lädt alle Listen und schreibt die zusammengeführte CSV- und Excel-Datei.
fn download(config: &AdmidioConfig, tx: &Sender<DownloadEvent>) -> Result<PathBuf> {
    if config.urls.is_empty() {
        bail!("Keine Export-URL angegeben");
    }
    if config.username.trim().is_empty() || config.password.is_empty() {
        bail!("Benutzername und Passwort sind erforderlich");
    }
    let renames = column_renames(&config.columns)?;

    let client = Client::builder()
        .cookie_store(true)
        .timeout(Duration::from_secs(180))
        .user_agent("admidio-serienbrief-client")
        .build()
        .context("HTTP-Client konnte nicht erstellt werden")?;

    let login = login_url(&config.urls[0])?;
    log(tx, format!("Anmeldung bei {login}"));
    let page = client
        .get(&login)
        .send()
        .and_then(|response| response.error_for_status())
        .with_context(|| format!("Anmeldeseite {login} nicht erreichbar"))?
        .text()?;
    let token = csrf_token(&page)
        .context("Kein CSRF-Token auf der Anmeldeseite gefunden, ist das eine Admidio-Installation?")?;

    let body = client
        .post(format!("{login}?mode=check"))
        .header("X-Requested-With", "XMLHttpRequest")
        .form(&[
            ("adm_csrf_token", token.as_str()),
            ("usr_login_name", config.username.trim()),
            ("usr_password", config.password.as_str()),
        ])
        .send()
        .and_then(|response| response.error_for_status())
        .context("Anmeldung konnte nicht gesendet werden")?
        .text()?;
    let answer: serde_json::Value =
        serde_json::from_str(&body).context("Antwort der Anmeldung ist kein JSON")?;
    if answer["status"] != "success" {
        let message = answer["message"].as_str().unwrap_or(body.as_str());
        bail!("Anmeldung fehlgeschlagen: {}", clean_message(message));
    }
    log(tx, "Anmeldung erfolgreich.");

    let mut table = Table::with_renames(renames);
    for url in &config.urls {
        let response = client
            .get(url)
            .send()
            .and_then(|response| response.error_for_status())
            .with_context(|| format!("Download von {url} fehlgeschlagen"))?;
        let content_type = header_value(&response, CONTENT_TYPE);
        let disposition = header_value(&response, CONTENT_DISPOSITION);
        let bytes = response.bytes()?;
        if !content_type.contains("csv") && !disposition.contains("attachment") {
            bail!(
                "{url} lieferte keine CSV-Datei (Content-Type '{content_type}'). \
                 Bitte URL und Berechtigung des Benutzers prüfen."
            );
        }
        let added = table
            .add_csv(&bytes)
            .with_context(|| format!("CSV von {url} ist ungültig"))?;
        log(tx, format!("{added} neue Datensätze von {url}"));
    }

    let renamed = table.applied_renames();
    if !renamed.is_empty() {
        log(tx, format!("Spalten umbenannt: {}", renamed.join(", ")));
    }
    let missing = table.missing_renames();
    if !missing.is_empty() {
        log(
            tx,
            format!(
                "Warnung: diese Admidio-Felder aus der Spaltenzuordnung kamen in keinem Export vor: {}",
                missing.join(", ")
            ),
        );
    }

    if let Some(otp) = &config.otp {
        let date = chrono::Local::now().format("%Y%m%d").to_string();
        let url = otp_url(otp, &date)?;
        log(tx, format!("Einmalpasswörter werden von {} abgerufen ...", otp.site_url.trim()));
        let body = client
            .get(&url)
            .send()
            .and_then(|response| response.error_for_status())
            .context("Abruf der OTP-Liste fehlgeschlagen")?
            .text()?;
        let entries = parse_otp_list(&body)?;
        let (matched, unmatched) = table.add_otps(&entries);
        log(
            tx,
            format!(
                "{} Einmalpasswörter erhalten, {matched} Datensätzen zugeordnet, Spalten {} ergänzt.",
                entries.len(),
                OTP_COLUMNS.join(", ")
            ),
        );
        if unmatched > 0 {
            log(tx, format!("Warnung: {unmatched} Datensätze ohne Einmalpasswort."));
        }
    }

    if let Some(parent) = config.target.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("Ordner '{}' kann nicht angelegt werden", parent.display()))?;
    }
    let xlsx_target = config.target.with_extension("xlsx");
    fs::write(&xlsx_target, table.to_xlsx()?)
        .with_context(|| format!("'{}' kann nicht geschrieben werden", xlsx_target.display()))?;
    fs::write(&config.target, table.to_csv()?)
        .with_context(|| format!("'{}' kann nicht geschrieben werden", config.target.display()))?;
    log(
        tx,
        format!(
            "{} Datensätze mit {} Spalten nach {} und {} geschrieben.",
            table.rows.len(),
            table.columns.len(),
            config.target.display(),
            xlsx_target.display()
        ),
    );
    Ok(config.target.clone())
}

fn header_value(response: &reqwest::blocking::Response, name: reqwest::header::HeaderName) -> String {
    response
        .headers()
        .get(name)
        .and_then(|value| value.to_str().ok())
        .unwrap_or("")
        .to_lowercase()
}

/// Sucht zum Wert der Spalte `column` eines Datensatzes den passenden OTP-Eintrag.
fn lookup_otp<'a>(
    row: &HashMap<String, String>,
    column: &str,
    map: &HashMap<&str, &'a OtpEntry>,
) -> Option<&'a OtpEntry> {
    let value = row.get(column)?.trim();
    if value.is_empty() {
        return None;
    }
    map.get(value).copied()
}

/// Zusammengeführte Tabelle mehrerer CSV-Exporte. Spalten werden nach Namen vereinigt,
/// vollständig identische Datensätze (Mitglied in mehreren Rollen) nur einmal übernommen.
/// Spaltennamen werden beim Einlesen gemäß der Umbenennungstabelle ersetzt.
#[derive(Default)]
pub struct Table {
    pub columns: Vec<String>,
    pub rows: Vec<HashMap<String, String>>,
    /// Umbenennungen Admidio-Feld → CSV-Spalte.
    renames: HashMap<String, String>,
    /// Originale Spaltennamen aller bisher eingelesenen Exporte.
    source_headers: HashSet<String>,
}

impl Table {
    /// Erzeugt eine leere Tabelle mit Umbenennungstabelle (Admidio-Feld → CSV-Spalte).
    pub fn with_renames(renames: HashMap<String, String>) -> Self {
        Self { renames, ..Self::default() }
    }

    /// Umbenennungen, die in mindestens einem Export angewendet wurden, als `alt → neu`.
    pub fn applied_renames(&self) -> Vec<String> {
        let mut applied: Vec<String> = self
            .renames
            .iter()
            .filter(|(admidio, _)| self.source_headers.contains(*admidio))
            .map(|(admidio, csv)| format!("{admidio} → {csv}"))
            .collect();
        applied.sort();
        applied
    }

    /// Admidio-Felder der Umbenennungstabelle, die in keinem Export vorkamen.
    pub fn missing_renames(&self) -> Vec<String> {
        let mut missing: Vec<String> = self
            .renames
            .keys()
            .filter(|admidio| !self.source_headers.contains(*admidio))
            .cloned()
            .collect();
        missing.sort();
        missing
    }

    /// Fügt den Inhalt einer CSV-Datei hinzu und liefert die Anzahl neu aufgenommener Datensätze.
    pub fn add_csv(&mut self, bytes: &[u8]) -> Result<usize> {
        let bytes = bytes.strip_prefix(b"\xEF\xBB\xBF").unwrap_or(bytes);
        let delimiter = detect_delimiter(bytes);

        let mut reader = csv::ReaderBuilder::new()
            .delimiter(delimiter)
            .flexible(true)
            .has_headers(true)
            .from_reader(bytes);

        let headers: Vec<String> = reader
            .byte_headers()
            .context("Kopfzeile fehlt")?
            .iter()
            .map(|field| decode_text(field).trim().to_string())
            .map(|header| {
                self.source_headers.insert(header.clone());
                self.renames.get(&header).cloned().unwrap_or(header)
            })
            .collect();
        if headers.is_empty() {
            bail!("Kopfzeile ist leer");
        }
        for header in &headers {
            if !self.columns.contains(header) {
                self.columns.push(header.clone());
            }
        }

        let mut added = 0;
        for record in reader.byte_records() {
            let record = record.context("Datensatz ist ungültig")?;
            if record.iter().all(|field| field.trim_ascii().is_empty()) {
                continue;
            }
            let row: HashMap<String, String> = headers
                .iter()
                .zip(record.iter())
                .map(|(name, value)| (name.clone(), decode_text(value)))
                .collect();
            if !self.rows.contains(&row) {
                self.rows.push(row);
                added += 1;
            }
        }
        Ok(added)
    }

    /// Ergänzt die Spalten `otp1`/`otp2`. Zugeordnet wird zuerst über den Anmeldenamen
    /// (Spalte `username`), sonst über die Mitgliedsnummer (Spalte `Mitgliedsnummer`).
    /// Liefert die Anzahl der Datensätze mit und ohne Einmalpasswort.
    pub fn add_otps(&mut self, entries: &[OtpEntry]) -> (usize, usize) {
        for column in OTP_COLUMNS {
            if !self.columns.iter().any(|name| name == column) {
                self.columns.push(column.to_string());
            }
        }
        let by_login: HashMap<&str, &OtpEntry> = entries
            .iter()
            .filter(|entry| !entry.login.is_empty())
            .map(|entry| (entry.login.as_str(), entry))
            .collect();
        let by_number: HashMap<&str, &OtpEntry> = entries
            .iter()
            .filter(|entry| !entry.member_number.is_empty())
            .map(|entry| (entry.member_number.as_str(), entry))
            .collect();

        let mut matched = 0;
        let mut unmatched = 0;
        for row in &mut self.rows {
            let entry = lookup_otp(row, USERNAME_COLUMN, &by_login)
                .or_else(|| lookup_otp(row, MEMBER_NUMBER_COLUMN, &by_number));
            let (otp1, otp2) = match entry {
                Some(entry) => {
                    matched += 1;
                    (entry.otp1.clone(), entry.otp2.clone())
                }
                None => {
                    unmatched += 1;
                    (String::new(), String::new())
                }
            };
            row.insert(OTP_COLUMNS[0].to_string(), otp1);
            row.insert(OTP_COLUMNS[1].to_string(), otp2);
        }
        (matched, unmatched)
    }

    /// Schreibt die Tabelle als UTF-8-Datei ohne BOM mit Semikolon als Trennzeichen.
    pub fn to_csv(&self) -> Result<Vec<u8>> {
        if self.columns.is_empty() {
            bail!("Keine Daten zum Schreiben");
        }
        let mut writer = csv::WriterBuilder::new()
            .delimiter(b';')
            .terminator(csv::Terminator::CRLF)
            .from_writer(Vec::new());
        writer.write_record(&self.columns)?;
        for row in &self.rows {
            writer.write_record(
                self.columns
                    .iter()
                    .map(|column| row.get(column).map(String::as_str).unwrap_or("")),
            )?;
        }
        Ok(writer.into_inner()?)
    }

    /// Schreibt die Tabelle als Excel-Arbeitsmappe mit einer Tabelle.
    pub fn to_xlsx(&self) -> Result<Vec<u8>> {
        if self.columns.is_empty() {
            bail!("Keine Daten zum Schreiben");
        }
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        for (column_index, column) in self.columns.iter().enumerate() {
            let column_index =
                u16::try_from(column_index).context("Zu viele Spalten für eine Excel-Tabelle")?;
            worksheet
                .write_string(0, column_index, column)
                .context("Excel-Kopfzeile konnte nicht geschrieben werden")?;
        }
        for (row_index, row) in self.rows.iter().enumerate() {
            let row_index =
                u32::try_from(row_index + 1).context("Zu viele Zeilen für eine Excel-Tabelle")?;
            for (column_index, column) in self.columns.iter().enumerate() {
                let column_index = u16::try_from(column_index)
                    .context("Zu viele Spalten für eine Excel-Tabelle")?;
                worksheet
                    .write_string(
                        row_index,
                        column_index,
                        row.get(column).map(String::as_str).unwrap_or(""),
                    )
                    .with_context(|| {
                        format!(
                            "Excel-Zelle in Zeile {}, Spalte {} konnte nicht geschrieben werden",
                            row_index + 1,
                            column_index + 1
                        )
                    })?;
            }
        }
        workbook
            .save_to_buffer()
            .context("Excel-Datei konnte nicht erstellt werden")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn login_url_is_derived_from_list_url() {
        let url = "https://verein.example.org/modules/groups-roles/lists_show.php?list_uuid=1&mode=csv";
        assert_eq!(login_url(url).unwrap(), "https://verein.example.org/system/login.php");
        let legacy = "https://verein.example.org/adm_program/modules/groups-roles/lists_show.php?mode=csv";
        assert_eq!(login_url(legacy).unwrap(), "https://verein.example.org/adm_program/system/login.php");
        assert!(login_url("https://verein.example.org/irgendwas.php").is_err());
    }

    #[test]
    fn csrf_token_is_extracted() {
        let html = r#"<form><input class="form-control" type="text" name="adm_csrf_token" id="adm_csrf_token" value="abc123XYZ"
            readonly><input name="usr_login_name" value=""></form>"#;
        assert_eq!(csrf_token(html).as_deref(), Some("abc123XYZ"));
        assert_eq!(csrf_token("<html></html>"), None);
    }

    #[test]
    fn error_message_is_cleaned() {
        let raw = "Der eingegebene Anmeldename und/oder das Passwort ist falsch. in /homepages/x/ModuleLogin.php, in line 175<br /><br />Stacktrace:<br />#0 ...";
        assert_eq!(clean_message(raw), "Der eingegebene Anmeldename und/oder das Passwort ist falsch.");
    }

    #[test]
    fn tables_are_merged_by_column_name_without_duplicates() {
        let mut table = Table::default();
        let first = table.add_csv(b"Mitgliedsnummer,username,Vorname\r\n1,max,Max\r\n2,erika,Erika\r\n").unwrap();
        let second = table.add_csv(b"Mitgliedsnummer,username,Ort\r\n2,erika,Ulm\r\n3,hans,Ehingen\r\n").unwrap();
        let third = table.add_csv(b"Mitgliedsnummer,username,Vorname\r\n1,max,Max\r\n").unwrap();
        assert_eq!((first, second, third), (2, 2, 0));
        assert_eq!(table.columns, vec!["Mitgliedsnummer", "username", "Vorname", "Ort"]);
        let bytes = table.to_csv().unwrap();
        assert!(!bytes.starts_with(&[0xEF, 0xBB, 0xBF]));
        let csv = String::from_utf8(bytes).unwrap();
        assert!(csv.starts_with("Mitgliedsnummer;username;Vorname;Ort\r\n"));
        assert!(csv.contains("3;hans;;Ehingen\r\n"));
        assert_eq!(csv.lines().count(), 5);

        let xlsx = table.to_xlsx().unwrap();
        assert!(xlsx.starts_with(b"PK\x03\x04"));
    }

    #[test]
    fn columns_are_renamed_on_import() {
        let mapping = vec![
            ColumnMapping { csv: "Mitgliedsnummer".into(), admidio: "Mitgliedsnr.".into() },
            ColumnMapping { csv: "username".into(), admidio: "Benutzername".into() },
            ColumnMapping { csv: "Fehlt".into(), admidio: "Nicht da".into() },
            ColumnMapping::default(),
        ];
        let mut table = Table::with_renames(column_renames(&mapping).unwrap());
        table.add_csv(b"Mitgliedsnr.,Benutzername,Vorname\r\n1,max,Max\r\n").unwrap();
        table.add_csv(b"Mitgliedsnummer,Benutzername,Ort\r\n1,max,Ulm\r\n").unwrap();
        assert_eq!(table.columns, vec!["Mitgliedsnummer", "username", "Vorname", "Ort"]);
        assert_eq!(table.rows[1]["username"], "max");
        assert_eq!(table.applied_renames(), vec!["Benutzername → username", "Mitgliedsnr. → Mitgliedsnummer"]);
        assert_eq!(table.missing_renames(), vec!["Nicht da"]);
    }

    #[test]
    fn column_mapping_is_validated() {
        let half = vec![ColumnMapping { csv: "a".into(), admidio: "".into() }];
        assert!(column_renames(&half).unwrap_err().to_string().contains("Zeile 1"));
        let duplicate_csv = vec![
            ColumnMapping { csv: "a".into(), admidio: "x".into() },
            ColumnMapping { csv: " a ".into(), admidio: "y".into() },
        ];
        assert!(column_renames(&duplicate_csv).unwrap_err().to_string().contains("CSV-Spalte 'a'"));
        let duplicate_admidio = vec![
            ColumnMapping { csv: "a".into(), admidio: "x".into() },
            ColumnMapping { csv: "b".into(), admidio: "x".into() },
        ];
        assert!(column_renames(&duplicate_admidio).unwrap_err().to_string().contains("Admidio-Feld 'x'"));
        assert!(column_renames(&[]).unwrap().is_empty());
    }

    #[test]
    fn otp_url_contains_md5_nonce_of_date() {
        let config = OtpConfig {
            site_url: "https://bfv-ehingen.de/".into(),
            password: "geheim".into(),
            clear: true,
        };
        // Nonce ist MD5("20240101")
        assert_eq!(
            otp_url(&config, "20240101").unwrap(),
            "https://bfv-ehingen.de/?getOTPlist=1&passwd=geheim&nonce=8696f8389f115a3bbb8208946b990cc1&clearotp=1"
        );
        let no_clear = OtpConfig { clear: false, ..config };
        assert!(otp_url(&no_clear, "20240101").unwrap().ends_with("&clearotp=0"));
        let invalid = OtpConfig { site_url: "bfv-ehingen.de".into(), ..no_clear };
        assert!(otp_url(&invalid, "20240101").is_err());
    }

    #[test]
    fn otp_list_is_parsed_and_matched() {
        let body = r#"{"BenKunter": ["BenKunter", "6", "Ohocb5B65nzbl", "zGI0MUSNvqfVe"],
                       "erika": ["erika", 2, "aaa", "bbb"]}"#;
        let entries = parse_otp_list(body).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(
            entries[0],
            OtpEntry {
                login: "BenKunter".into(),
                member_number: "6".into(),
                otp1: "Ohocb5B65nzbl".into(),
                otp2: "zGI0MUSNvqfVe".into()
            }
        );
        assert_eq!(entries[1].member_number, "2");

        let mut table = Table::default();
        table
            .add_csv(b"Mitgliedsnummer,username,Vorname\r\n6,BenKunter,Ben\r\n2,unbekannt,Erika\r\n9,niemand,Hans\r\n")
            .unwrap();
        assert_eq!(table.add_otps(&entries), (2, 1));
        assert_eq!(table.columns, vec!["Mitgliedsnummer", "username", "Vorname", "otp1", "otp2"]);
        assert_eq!(table.rows[0]["otp1"], "Ohocb5B65nzbl");
        assert_eq!(table.rows[1]["otp2"], "bbb");
        assert_eq!(table.rows[2]["otp1"], "");

        assert!(parse_otp_list("Anfrage gesperrt").unwrap_err().to_string().contains("kein JSON"));
        assert!(parse_otp_list(r#"{"x": ["x", "1"]}"#).unwrap_err().to_string().contains("4 Felder"));
    }

    #[test]
    fn role_list_is_replaced_in_url() {
        let url = "https://verein.example.org/modules/groups-roles/lists_show.php?list_uuid=abc&role_list=alt&date_from=2026-01-01&date_to=2026-12-31&mode=csv";
        assert_eq!(
            with_query_params(url, &[("role_list", " 11111111-2222-3333-4444-555555555555 ")]).unwrap(),
            "https://verein.example.org/modules/groups-roles/lists_show.php?list_uuid=abc&date_from=2026-01-01&date_to=2026-12-31&mode=csv&role_list=11111111-2222-3333-4444-555555555555"
        );
        // Ohne vorhandenen Parameter wird er ergänzt; Kommas bleiben URL-kodiert erhalten.
        let without = "https://verein.example.org/modules/groups-roles/lists_show.php?list_uuid=abc&mode=csv";
        assert_eq!(
            with_query_params(without, &[("role_list", "a,b")]).unwrap(),
            "https://verein.example.org/modules/groups-roles/lists_show.php?list_uuid=abc&mode=csv&role_list=a%2Cb"
        );
        assert!(with_query_params("kein://[url", &[("role_list", "a")]).is_err());
    }

    #[test]
    fn role_list_is_replaced_in_every_url_line() {
        let text = "# Kommentar\nhttps://a/modules/x?list_uuid=1&role_list=alt&mode=csv\n\n  https://b/modules/y?mode=csv  ";
        assert_eq!(
            replace_role_list(text, "neu").unwrap(),
            "# Kommentar\nhttps://a/modules/x?list_uuid=1&mode=csv&role_list=neu\n\nhttps://b/modules/y?mode=csv&role_list=neu"
        );
        assert!(replace_role_list("kein://[url", "a").is_err());
    }

    #[test]
    fn dates_are_replaced_and_read_back() {
        let text = "https://a/modules/x?list_uuid=1&date_from=2020-01-01&role_list=r&date_to=2020-12-31&mode=csv";
        let today = [("date_from", "2026-10-23"), ("date_to", "2026-10-23")];
        let replaced = replace_query_params(text, &today).unwrap();
        assert_eq!(
            replaced,
            "https://a/modules/x?list_uuid=1&role_list=r&mode=csv&date_from=2026-10-23&date_to=2026-10-23"
        );
        assert_eq!(query_param(&replaced, "date_from").as_deref(), Some("2026-10-23"));
        assert_eq!(query_param(&replaced, "role_list").as_deref(), Some("r"));
        assert_eq!(query_param("# nur Kommentar\n", "role_list"), None);
        assert_eq!(query_param("https://a/modules/x?mode=csv", "role_list"), None);
        assert_eq!(query_param("kein://[url", "mode"), None);
    }

    #[test]
    fn role_buttons_are_validated() {
        let buttons = vec![
            RoleButton { name: " Vorstand ".into(), role_list: " uuid-1 ".into() },
            RoleButton::default(),
            RoleButton { name: "Alle".into(), role_list: "uuid-1,uuid-2".into() },
        ];
        let valid = role_buttons(&buttons).unwrap();
        assert_eq!(valid.len(), 2);
        assert_eq!(valid[0], RoleButton { name: "Vorstand".into(), role_list: "uuid-1".into() });

        let half = vec![RoleButton { name: "Vorstand".into(), role_list: "".into() }];
        assert!(role_buttons(&half).unwrap_err().to_string().contains("Zeile 1"));
        let duplicate = vec![
            RoleButton { name: "Vorstand".into(), role_list: "a".into() },
            RoleButton { name: "Vorstand".into(), role_list: "b".into() },
        ];
        assert!(role_buttons(&duplicate).unwrap_err().to_string().contains("'Vorstand'"));
    }

    #[test]
    fn urls_are_parsed_per_line() {
        let urls = parse_urls("  https://a/modules/x?mode=csv \n\n# Kommentar\nhttps://b/modules/y?mode=csv\n");
        assert_eq!(urls, vec!["https://a/modules/x?mode=csv", "https://b/modules/y?mode=csv"]);
    }

    /// Prüft Anmeldeseite und Fehlerbehandlung gegen eine echte Installation. Erwartet in
    /// `ADMIDIO_LIST_URL` eine Listen-URL; die Anmeldung mit falschem Passwort muss sauber
    /// abgelehnt werden. Start mit `cargo test admidio_login -- --ignored --nocapture`.
    #[test]
    #[ignore]
    fn admidio_login_rejects_wrong_password() {
        let url = std::env::var("ADMIDIO_LIST_URL").unwrap();
        let config = AdmidioConfig {
            username: "serienbrief-test-unbekannt".to_string(),
            password: "falsch".to_string(),
            urls: vec![url],
            target: std::env::temp_dir().join("serienbrief_login_test.csv"),
            columns: Vec::new(),
            otp: None,
        };
        let (tx, rx) = std::sync::mpsc::channel();
        let error = download(&config, &tx).unwrap_err();
        for event in rx.try_iter() {
            println!("{event:?}");
        }
        let message = format!("{error:#}");
        println!("{message}");
        assert!(message.starts_with("Anmeldung fehlgeschlagen: "));
        assert!(!message.contains("Stacktrace"));
    }
}

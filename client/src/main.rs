//! Grafische Oberfläche „Word Serienbrief“: erzeugt aus einer Word-Vorlage und einer
//! CSV-Datenquelle pro Datensatz eine PDF-Datei. Nachbau von `com_mail_merge_class.py`.

#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod admidio;
mod merge;
mod word;

use std::path::{Path, PathBuf};
use std::sync::mpsc::Receiver;
use std::time::{Duration, Instant};

use eframe::egui;

use admidio::{AdmidioConfig, ColumnMapping, DownloadEvent, OtpConfig};
use merge::{DataSourceInfo, MergeConfig, Progress};

/// Standard-Muster für die Dateinamen, entspricht dem Python-Original.
const DEFAULT_PATTERN: &str = "rundschreiben-{Mitgliedsnummer}-{username}.pdf";

/// Maximal angezeigte Protokollzeilen.
const MAX_LOG_LINES: usize = 2000;

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([820.0, 640.0])
            .with_min_inner_size([640.0, 480.0])
            .with_title("Word Serienbrief"),
        ..Default::default()
    };
    eframe::run_native(
        "admidio-serienbrief-client",
        options,
        Box::new(|cc| {
            let settings = cc
                .storage
                .and_then(|storage| eframe::get_value(storage, eframe::APP_KEY))
                .unwrap_or_default();
            let mut app = SerienbriefApp { settings, ..Default::default() };
            app.reload_data_source();
            Ok(Box::new(app))
        }),
    )
}

/// Eingaben, die zwischen zwei Programmstarts gespeichert werden.
#[derive(serde::Serialize, serde::Deserialize)]
#[serde(default)]
struct Settings {
    template: String,
    data_source: String,
    output_dir: String,
    file_pattern: String,
    thread_count: usize,
    kill_word: bool,
    /// Anmeldename für Admidio. Das Passwort wird bewusst nicht gespeichert.
    admidio_username: String,
    /// CSV-Export-URLs aus Admidio, eine je Zeile.
    admidio_urls: String,
    /// Zuordnung der CSV-Spalten zu den Feldnamen im Admidio-Export.
    admidio_columns: Vec<ColumnMapping>,
    /// Adresse der WordPress-Seite mit dem OTP-Endpunkt; leer schaltet den Abruf ab.
    otp_site_url: String,
    /// Passwort des OTP-Endpunkts (wird gespeichert, da es ein fester Dienstschlüssel ist).
    otp_password: String,
    /// Bestehende Einmalpasswörter auf dem Server löschen und neu erzeugen.
    otp_clear: bool,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            template: String::new(),
            data_source: String::new(),
            output_dir: String::new(),
            file_pattern: DEFAULT_PATTERN.to_string(),
            thread_count: 8,
            kill_word: false,
            admidio_username: String::new(),
            admidio_urls: String::new(),
            admidio_columns: Vec::new(),
            otp_site_url: String::new(),
            otp_password: String::new(),
            otp_clear: true,
        }
    }
}

/// Laufzeitzustand der Oberfläche.
#[derive(Default)]
struct SerienbriefApp {
    settings: Settings,
    data_info: Option<DataSourceInfo>,
    data_error: Option<String>,
    log: Vec<String>,
    receiver: Option<Receiver<Progress>>,
    total: usize,
    done: usize,
    failed: usize,
    started: Option<Instant>,
    finished: Option<Duration>,
    /// Admidio-Passwort, nur für die laufende Sitzung.
    admidio_password: String,
    /// Kanal des laufenden Downloads aus Admidio.
    download_receiver: Option<Receiver<DownloadEvent>>,
    /// Ergebnis des letzten Downloads für die Anzeige.
    download_status: Option<Result<String, String>>,
    /// Fenster mit der Spaltenzuordnung ist geöffnet.
    show_column_config: bool,
}

impl SerienbriefApp {
    /// Liest die Datenquelle neu ein und aktualisiert Spaltenliste und Vorschau.
    fn reload_data_source(&mut self) {
        self.data_info = None;
        self.data_error = None;
        if self.settings.data_source.trim().is_empty() {
            return;
        }
        match merge::inspect_data_source(Path::new(self.settings.data_source.trim())) {
            Ok(info) => self.data_info = Some(info),
            Err(error) => self.data_error = Some(format!("{error:#}")),
        }
    }

    /// Platzhalter des Musters, die in der Datenquelle nicht vorkommen.
    fn unknown_placeholders(&self) -> Vec<String> {
        let Some(info) = &self.data_info else {
            return Vec::new();
        };
        merge::placeholders(&self.settings.file_pattern)
            .into_iter()
            .filter(|name| !info.fields.contains(name))
            .collect()
    }

    /// Vorschlag für den Zielordner: `<Ordner der Vorlage>/<Datum>-Rundschreiben`.
    fn default_output_dir(template: &str) -> String {
        let folder = Path::new(template).parent().unwrap_or(Path::new("."));
        let name = chrono::Local::now().format("%Y-%m-%d-Rundschreiben").to_string();
        folder.join(name).to_string_lossy().into_owned()
    }

    fn is_running(&self) -> bool {
        self.receiver.is_some()
    }

    fn is_downloading(&self) -> bool {
        self.download_receiver.is_some()
    }

    /// Prüft die Admidio-Eingaben und liefert die Download-Konfiguration oder den ersten Fehler.
    fn validate_download(&self) -> Result<AdmidioConfig, String> {
        if self.settings.admidio_username.trim().is_empty() {
            return Err("Bitte den Admidio-Benutzernamen eingeben.".to_string());
        }
        if self.admidio_password.is_empty() {
            return Err("Bitte das Admidio-Passwort eingeben.".to_string());
        }
        let urls = admidio::parse_urls(&self.settings.admidio_urls);
        if urls.is_empty() {
            return Err("Bitte mindestens eine Export-URL eingeben.".to_string());
        }
        for url in &urls {
            if !url.starts_with("https://") && !url.starts_with("http://") {
                return Err(format!("Keine gültige URL: {url}"));
            }
            admidio::login_url(url).map_err(|error| error.to_string())?;
        }
        admidio::column_renames(&self.settings.admidio_columns).map_err(|error| error.to_string())?;
        let otp = self.validate_otp()?;
        Ok(AdmidioConfig {
            username: self.settings.admidio_username.trim().to_string(),
            password: self.admidio_password.clone(),
            urls,
            target: PathBuf::new(),
            columns: self.settings.admidio_columns.clone(),
            otp,
        })
    }

    /// Prüft die OTP-Eingaben; ohne WordPress-Adresse ist der Abruf abgeschaltet.
    fn validate_otp(&self) -> Result<Option<OtpConfig>, String> {
        let site_url = self.settings.otp_site_url.trim();
        if site_url.is_empty() {
            return Ok(None);
        }
        if !site_url.starts_with("https://") && !site_url.starts_with("http://") {
            return Err(format!("Keine gültige WordPress-Adresse: {site_url}"));
        }
        if self.settings.otp_password.trim().is_empty() {
            return Err("Bitte das Passwort des OTP-Endpunkts eingeben.".to_string());
        }
        Ok(Some(OtpConfig {
            site_url: site_url.to_string(),
            password: self.settings.otp_password.trim().to_string(),
            clear: self.settings.otp_clear,
        }))
    }

    /// Anzahl der ausgefüllten Zeilen der Spaltenzuordnung.
    fn column_mapping_count(&self) -> usize {
        self.settings
            .admidio_columns
            .iter()
            .filter(|mapping| !mapping.csv.trim().is_empty() || !mapping.admidio.trim().is_empty())
            .count()
    }

    /// Zeichnet das Fenster mit der Spaltenzuordnung (CSV-Spalte ↔ Feld im Admidio-Export).
    fn column_config_window(&mut self, ctx: &egui::Context, enabled: bool) {
        let mut open = self.show_column_config;
        egui::Window::new("Spaltenzuordnung für den Admidio-Download")
            .open(&mut open)
            .default_width(600.0)
            .resizable(true)
            .collapsible(false)
            .show(ctx, |ui| {
                ui.label(
                    "Links der Spaltenname, den der Serienbrief (CSV) erwartet, rechts der Name des \
                     Felds im Admidio-Export. Beim Download werden die Spalten entsprechend \
                     umbenannt. Spalten, die hier nicht aufgeführt sind, behalten ihren Namen aus \
                     Admidio.",
                );
                ui.add_space(6.0);

                let mut remove = None;
                egui::ScrollArea::vertical().max_height(360.0).show(ui, |ui| {
                    egui::Grid::new("column_mapping")
                        .num_columns(3)
                        .spacing([12.0, 6.0])
                        .striped(true)
                        .show(ui, |ui| {
                            ui.strong("Spalte in CSV / Serienbrief");
                            ui.strong("Feld im Admidio-Export");
                            ui.label("");
                            ui.end_row();
                            for (index, mapping) in self.settings.admidio_columns.iter_mut().enumerate() {
                                ui.add_enabled(
                                    enabled,
                                    egui::TextEdit::singleline(&mut mapping.csv).desired_width(240.0),
                                );
                                ui.add_enabled(
                                    enabled,
                                    egui::TextEdit::singleline(&mut mapping.admidio).desired_width(240.0),
                                );
                                if ui
                                    .add_enabled(enabled, egui::Button::new("✖"))
                                    .on_hover_text("Zeile entfernen")
                                    .clicked()
                                {
                                    remove = Some(index);
                                }
                                ui.end_row();
                            }
                        });
                });
                if let Some(index) = remove {
                    self.settings.admidio_columns.remove(index);
                }

                ui.add_space(6.0);
                ui.horizontal(|ui| {
                    if ui.add_enabled(enabled, egui::Button::new("Zeile hinzufügen")).clicked() {
                        self.settings.admidio_columns.push(ColumnMapping::default());
                    }
                    let fields = self
                        .data_info
                        .as_ref()
                        .map(|info| info.fields.clone())
                        .unwrap_or_default();
                    if ui
                        .add_enabled(
                            enabled && !fields.is_empty(),
                            egui::Button::new("Spalten der Datenquelle übernehmen"),
                        )
                        .on_hover_text(
                            "Legt für jede Spalte der aktuellen Datenquelle eine Zeile an, die noch fehlt. \
                             Das Admidio-Feld wird mit demselben Namen vorbelegt.",
                        )
                        .clicked()
                    {
                        for field in fields {
                            let exists = self
                                .settings
                                .admidio_columns
                                .iter()
                                .any(|mapping| mapping.csv.trim() == field);
                            if !exists {
                                self.settings.admidio_columns.push(ColumnMapping {
                                    csv: field.clone(),
                                    admidio: field,
                                });
                            }
                        }
                    }
                    if ui
                        .add_enabled(
                            enabled && !self.settings.admidio_columns.is_empty(),
                            egui::Button::new("Alle entfernen"),
                        )
                        .clicked()
                    {
                        self.settings.admidio_columns.clear();
                    }
                });

                ui.add_space(4.0);
                match admidio::column_renames(&self.settings.admidio_columns) {
                    Ok(renames) => {
                        ui.label(format!("{} Spalten werden beim Download umbenannt.", renames.len()));
                    }
                    Err(error) => {
                        ui.colored_label(egui::Color32::RED, error.to_string());
                    }
                }
            });
        self.show_column_config = open;
    }

    /// Fragt bei leerer Datenquelle nach einer Zieldatei und startet den Download.
    fn start_download(&mut self, mut config: AdmidioConfig) {
        let current = self.settings.data_source.trim();
        let target = if current.is_empty() {
            let mut dialog = rfd::FileDialog::new()
                .add_filter("CSV-Dateien", &["csv"])
                .set_file_name("admidio_export.csv");
            if let Some(folder) = Path::new(self.settings.template.trim()).parent() {
                if folder.is_dir() {
                    dialog = dialog.set_directory(folder);
                }
            }
            match dialog.save_file() {
                Some(path) => path,
                None => return,
            }
        } else {
            PathBuf::from(current)
        };
        self.settings.data_source = target.to_string_lossy().into_owned();
        config.target = target;

        let (tx, rx) = std::sync::mpsc::channel();
        self.download_status = None;
        self.download_receiver = Some(rx);
        self.log.push("Datenquelle wird aus Admidio aktualisiert ...".to_string());
        admidio::start(config, tx);
    }

    /// Holt die Meldungen des Download-Threads ab.
    fn poll_download(&mut self) {
        let Some(receiver) = &self.download_receiver else {
            return;
        };
        let mut finished = None;
        for event in receiver.try_iter() {
            match event {
                DownloadEvent::Log(line) => self.log.push(line),
                DownloadEvent::Finished(result) => finished = Some(result),
            }
        }
        if let Some(result) = finished {
            self.download_receiver = None;
            self.download_status = Some(match result {
                Ok(path) => {
                    self.reload_data_source();
                    Ok(format!("Datenquelle aktualisiert: {}", path.display()))
                }
                Err(message) => {
                    self.log.push(message.clone());
                    Err(message)
                }
            });
        }
    }

    /// Zeichnet den aufklappbaren Bereich zum Aktualisieren der Datenquelle aus Admidio.
    fn admidio_section(&mut self, ui: &mut egui::Ui, editable: bool) {
        egui::CollapsingHeader::new("Datenquelle aus Admidio aktualisieren")
            .default_open(!self.settings.admidio_urls.trim().is_empty())
            .show(ui, |ui| {
                let enabled = editable && !self.is_downloading();
                egui::Grid::new("admidio")
                    .num_columns(2)
                    .spacing([12.0, 8.0])
                    .show(ui, |ui| {
                        ui.label("Benutzername");
                        ui.add_enabled(
                            enabled,
                            egui::TextEdit::singleline(&mut self.settings.admidio_username)
                                .desired_width(f32::INFINITY),
                        );
                        ui.end_row();

                        ui.label("Passwort");
                        ui.add_enabled(
                            enabled,
                            egui::TextEdit::singleline(&mut self.admidio_password)
                                .password(true)
                                .desired_width(f32::INFINITY),
                        );
                        ui.end_row();

                        ui.label("Export-URLs\n(eine je Rolle)");
                        ui.add_enabled(
                            enabled,
                            egui::TextEdit::multiline(&mut self.settings.admidio_urls)
                                .desired_rows(3)
                                .desired_width(f32::INFINITY)
                                .hint_text(
                                    "https://.../modules/groups-roles/lists_show.php?list_uuid=...&role_list=...&mode=csv",
                                ),
                        );
                        ui.end_row();

                        ui.label("Spaltenzuordnung");
                        ui.horizontal(|ui| {
                            if ui.button("Bearbeiten …").clicked() {
                                self.show_column_config = true;
                            }
                            let count = self.column_mapping_count();
                            if count == 0 {
                                ui.label("keine, Spaltennamen werden unverändert übernommen");
                            } else {
                                ui.label(format!("{count} Spalten zugeordnet"));
                            }
                        });
                        ui.end_row();

                        ui.label("Einmalpasswörter\n(WordPress)");
                        ui.add_enabled(
                            enabled,
                            egui::TextEdit::singleline(&mut self.settings.otp_site_url)
                                .desired_width(f32::INFINITY)
                                .hint_text("https://<wordpress-seite>/ – leer lassen, um keine OTP-Spalten zu ergänzen"),
                        );
                        ui.end_row();

                        ui.label("OTP-Passwort");
                        ui.add_enabled(
                            enabled,
                            egui::TextEdit::singleline(&mut self.settings.otp_password)
                                .password(true)
                                .desired_width(f32::INFINITY),
                        );
                        ui.end_row();

                        ui.label("");
                        ui.add_enabled(
                            enabled,
                            egui::Checkbox::new(
                                &mut self.settings.otp_clear,
                                "Einmalpasswörter auf dem Server neu erzeugen (clearotp=1)",
                            ),
                        );
                        ui.end_row();
                    });

                let validation = self.validate_download();
                ui.horizontal(|ui| {
                    let can_start = enabled && validation.is_ok();
                    if ui
                        .add_enabled(can_start, egui::Button::new("CSV aus Admidio laden"))
                        .clicked()
                    {
                        if let Ok(config) = &validation {
                            self.start_download(config.clone());
                        }
                    }
                    if self.is_downloading() {
                        ui.spinner();
                        ui.label("Download läuft ...");
                    } else if let Err(message) = &validation {
                        ui.colored_label(egui::Color32::GRAY, message);
                    } else if let Some(status) = &self.download_status {
                        match status {
                            Ok(message) => ui.colored_label(egui::Color32::DARK_GREEN, message),
                            Err(message) => ui.colored_label(egui::Color32::RED, message),
                        };
                    }
                });
                ui.label(
                    "Die Exporte werden nach Spaltennamen zusammengeführt, gemäß Spaltenzuordnung \
                     umbenannt und in die oben gewählte Datenquelle geschrieben. Das Admidio-Passwort \
                     wird nicht gespeichert. Ist eine WordPress-Adresse eingetragen, werden die \
                     Einmalpasswörter abgerufen, über die Spalten username bzw. Mitgliedsnummer \
                     zugeordnet und als Spalten otp1 und otp2 ergänzt.",
                );
            });
    }

    /// Prüft die Eingaben und liefert die Konfiguration oder den ersten Fehler.
    fn validate(&self) -> Result<MergeConfig, String> {
        let template = PathBuf::from(self.settings.template.trim());
        if !template.is_file() {
            return Err("Bitte eine vorhandene Word-Vorlage auswählen.".to_string());
        }
        let data_source = PathBuf::from(self.settings.data_source.trim());
        if !data_source.is_file() {
            return Err("Bitte eine vorhandene Datenquelle (CSV) auswählen.".to_string());
        }
        if let Some(error) = &self.data_error {
            return Err(error.clone());
        }
        let output_dir = self.settings.output_dir.trim();
        if output_dir.is_empty() {
            return Err("Bitte einen Zielordner angeben.".to_string());
        }
        let pattern = self.settings.file_pattern.trim();
        if pattern.is_empty() || merge::placeholders(pattern).is_empty() {
            return Err("Das Dateinamen-Muster braucht mindestens einen Platzhalter.".to_string());
        }
        let unknown = self.unknown_placeholders();
        if !unknown.is_empty() {
            return Err(format!("Unbekannte Platzhalter: {}", unknown.join(", ")));
        }
        Ok(MergeConfig {
            template,
            data_source,
            output_dir: PathBuf::from(output_dir),
            file_pattern: pattern.to_string(),
            thread_count: self.settings.thread_count.clamp(1, 32),
            kill_word: self.settings.kill_word,
        })
    }

    /// Startet den Serienbrief-Lauf.
    fn start(&mut self, config: MergeConfig) {
        let (tx, rx) = std::sync::mpsc::channel();
        self.log.clear();
        self.total = 0;
        self.done = 0;
        self.failed = 0;
        self.finished = None;
        self.started = Some(Instant::now());
        self.receiver = Some(rx);
        merge::start(config, tx);
    }

    /// Holt alle anstehenden Fortschrittsmeldungen ab.
    fn poll_progress(&mut self) {
        let Some(receiver) = &self.receiver else {
            return;
        };
        let mut finished = None;
        for message in receiver.try_iter() {
            match message {
                Progress::Total(total) => self.total = total,
                Progress::Log(line) => {
                    self.log.push(line);
                    if self.log.len() > MAX_LOG_LINES {
                        let overflow = self.log.len() - MAX_LOG_LINES;
                        self.log.drain(..overflow);
                    }
                }
                Progress::RecordDone => self.done += 1,
                Progress::RecordFailed => self.failed += 1,
                Progress::Finished(duration) => finished = Some(duration),
            }
        }
        if let Some(duration) = finished {
            self.finished = Some(duration);
            self.receiver = None;
            self.log.push(format!(
                "Fertig: {} PDF-Dateien, {} Fehler, Gesamtdauer {}",
                self.done,
                self.failed,
                merge::format_duration(duration)
            ));
        }
    }

    /// Zeichnet eine Zeile mit Beschriftung, Textfeld und Auswahl-Knopf; liefert `true`,
    /// wenn sich der Wert geändert hat.
    fn path_row(
        ui: &mut egui::Ui,
        label: &str,
        value: &mut String,
        enabled: bool,
        pick: impl FnOnce() -> Option<PathBuf>,
    ) -> bool {
        let mut changed = false;
        ui.label(label);
        ui.add_enabled_ui(enabled, |ui| {
            // Knopf zuerst (rechtsbündig), das Textfeld füllt den restlichen Platz.
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if ui.button("…").clicked() {
                    if let Some(path) = pick() {
                        *value = path.to_string_lossy().into_owned();
                        changed = true;
                    }
                }
                if ui
                    .add(egui::TextEdit::singleline(value).desired_width(f32::INFINITY))
                    .changed()
                {
                    changed = true;
                }
            });
        });
        ui.end_row();
        changed
    }

    /// Vorschau des Dateinamens mit den Werten des ersten Datensatzes.
    fn preview(&self) -> Option<String> {
        let info = self.data_info.as_ref()?;
        if info.first_record.is_empty() {
            return None;
        }
        Some(merge::render(&self.settings.file_pattern, &info.first_record))
    }

    /// Zeichnet Fortschrittsbalken mit vergangener und geschätzter Restzeit.
    fn progress_bar(&self, ui: &mut egui::Ui) {
        let processed = self.done + self.failed;
        let fraction = if self.total > 0 {
            processed as f32 / self.total as f32
        } else {
            0.0
        };
        let elapsed = match (self.finished, self.started) {
            (Some(duration), _) => duration,
            (None, Some(started)) => started.elapsed(),
            _ => Duration::ZERO,
        };
        let remaining = if processed > 0 && self.total > processed {
            let per_record = elapsed.as_secs_f64() / processed as f64;
            Duration::from_secs_f64(per_record * (self.total - processed) as f64)
        } else {
            Duration::ZERO
        };
        let text = format!(
            "{processed} / {} | {} vergangen | ca. {} verbleibend",
            self.total,
            merge::format_duration(elapsed),
            merge::format_duration(remaining)
        );
        ui.add(egui::ProgressBar::new(fraction).text(text).animate(self.is_running()));
    }
}

impl eframe::App for SerienbriefApp {
    fn save(&mut self, storage: &mut dyn eframe::Storage) {
        eframe::set_value(storage, eframe::APP_KEY, &self.settings);
    }

    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.poll_progress();
        self.poll_download();
        if self.is_running() || self.is_downloading() {
            ctx.request_repaint_after(Duration::from_millis(200));
        }
        let editable = !self.is_running();

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Word Serienbrief als PDF");
            ui.add_space(8.0);

            let mut data_changed = false;
            let mut template_changed = false;

            egui::Grid::new("settings")
                .num_columns(2)
                .spacing([12.0, 8.0])
                .show(ui, |ui| {
                    template_changed = Self::path_row(
                        ui,
                        "Serienbrief (Word)",
                        &mut self.settings.template,
                        editable,
                        || {
                            rfd::FileDialog::new()
                                .add_filter("Word-Dokumente", &["docx", "docm", "doc", "dotx", "dotm"])
                                .add_filter("Alle Dateien", &["*"])
                                .pick_file()
                        },
                    );
                    data_changed = Self::path_row(
                        ui,
                        "Datenquelle (CSV)",
                        &mut self.settings.data_source,
                        editable,
                        || {
                            rfd::FileDialog::new()
                                .add_filter("CSV-Dateien", &["csv", "txt"])
                                .add_filter("Alle Dateien", &["*"])
                                .pick_file()
                        },
                    );
                    Self::path_row(
                        ui,
                        "Zielordner",
                        &mut self.settings.output_dir,
                        editable,
                        || rfd::FileDialog::new().pick_folder(),
                    );

                    ui.label("Dateiname");
                    ui.add_enabled_ui(editable, |ui| {
                        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                            if ui.button("Standard").clicked() {
                                self.settings.file_pattern = DEFAULT_PATTERN.to_string();
                            }
                            ui.add(
                                egui::TextEdit::singleline(&mut self.settings.file_pattern)
                                    .desired_width(f32::INFINITY),
                            );
                        });
                    });
                    ui.end_row();

                    ui.label("Parallele Word-Instanzen");
                    ui.add_enabled(
                        editable,
                        egui::Slider::new(&mut self.settings.thread_count, 1..=16),
                    );
                    ui.end_row();

                    ui.label("");
                    ui.add_enabled(
                        editable,
                        egui::Checkbox::new(
                            &mut self.settings.kill_word,
                            "Vor dem Start alle Word-Prozesse beenden (taskkill)",
                        ),
                    );
                    ui.end_row();
                });

            if template_changed && self.settings.output_dir.trim().is_empty() {
                self.settings.output_dir = Self::default_output_dir(&self.settings.template);
            }
            if data_changed {
                self.reload_data_source();
            }

            ui.add_space(4.0);
            self.admidio_section(ui, editable);
            ui.add_space(6.0);

            // Informationen zur Datenquelle und Vorschau des Dateinamens.
            if let Some(error) = &self.data_error {
                ui.colored_label(egui::Color32::RED, error);
            } else if let Some(info) = &self.data_info {
                ui.label(format!(
                    "{} Datensätze, Spalten: {}",
                    info.record_count,
                    info.fields.join(", ")
                ));
                let unknown = self.unknown_placeholders();
                if !unknown.is_empty() {
                    ui.colored_label(
                        egui::Color32::RED,
                        format!("Unbekannte Platzhalter: {}", unknown.join(", ")),
                    );
                } else if let Some(preview) = self.preview() {
                    ui.label(format!("Vorschau (1. Datensatz): {preview}"));
                }
            }

            ui.add_space(8.0);

            // Start-Knopf und Fortschritt.
            let validation = self.validate();
            ui.horizontal(|ui| {
                let can_start = editable && !self.is_downloading() && validation.is_ok();
                if ui.add_enabled(can_start, egui::Button::new("PDFs erzeugen")).clicked() {
                    if let Ok(config) = &validation {
                        self.start(config.clone());
                    }
                }
                if let (false, Err(message)) = (self.is_running(), &validation) {
                    ui.colored_label(egui::Color32::GRAY, message);
                }
                if self.is_running() {
                    ui.spinner();
                }
            });

            if self.started.is_some() {
                self.progress_bar(ui);
            }

            ui.add_space(6.0);
            ui.separator();
            ui.label("Protokoll");
            egui::ScrollArea::vertical()
                .stick_to_bottom(true)
                .auto_shrink([false, false])
                .show(ui, |ui| {
                    for line in &self.log {
                        ui.monospace(line);
                    }
                });
        });

        self.column_config_window(ctx, editable && !self.is_downloading());
    }
}

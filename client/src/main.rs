//! Grafische Oberfläche „Word Serienbrief“: erzeugt aus einer Word-Vorlage und einer
//! CSV-Datenquelle pro Datensatz eine PDF-Datei. Nachbau von `com_mail_merge_class.py`.

#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod merge;
mod word;

use std::path::{Path, PathBuf};
use std::sync::mpsc::Receiver;
use std::time::{Duration, Instant};

use eframe::egui;

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
        if self.is_running() {
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
                let can_start = editable && validation.is_ok();
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
    }
}

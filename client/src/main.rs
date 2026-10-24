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

use admidio::{AdmidioConfig, ColumnMapping, DownloadEvent, OtpConfig, RoleButton};
use merge::{DataSourceInfo, MergeConfig, Progress};

/// Standard-Muster für die Dateinamen, entspricht dem Python-Original.
const DEFAULT_PATTERN: &str = "rundschreiben-{Mitgliedsnummer}-{username}.pdf";

/// Maximal angezeigte Protokollzeilen.
const MAX_LOG_LINES: usize = 2000;

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([860.0, 820.0])
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
    flyer_path: String,
    file_pattern: String,
    thread_count: usize,
    kill_word: bool,
    /// Anmeldename für Admidio. Das Passwort wird bewusst nicht gespeichert.
    admidio_username: String,
    /// CSV-Export-URLs aus Admidio, eine je Zeile.
    admidio_urls: String,
    /// Zuordnung der CSV-Spalten zu den Feldnamen im Admidio-Export.
    admidio_columns: Vec<ColumnMapping>,
    /// Konfigurierbare Knöpfe: Rollenname → Wert für den URL-Parameter `role_list`.
    admidio_roles: Vec<RoleButton>,
    /// Adresse der WordPress-Seite mit dem OTP-Endpunkt; leer schaltet den Abruf ab.
    otp_site_url: String,
    /// Passwort des OTP-Endpunkts (wird gespeichert, da es ein fester Dienstschlüssel ist).
    otp_password: String,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            template: String::new(),
            data_source: String::new(),
            output_dir: String::new(),
            flyer_path: String::new(),
            file_pattern: DEFAULT_PATTERN.to_string(),
            thread_count: 8,
            kill_word: false,
            admidio_username: String::new(),
            admidio_urls: String::new(),
            admidio_columns: Vec::new(),
            admidio_roles: Vec::new(),
            otp_site_url: String::new(),
            otp_password: String::new(),
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
    /// Einmalpasswörter beim nächsten Download von WordPress abrufen. Wird bewusst nicht
    /// gespeichert, damit der Abruf bei jedem Programmstart ausdrücklich eingeschaltet wird.
    otp_enabled: bool,
    /// Kanal des laufenden Downloads aus Admidio.
    download_receiver: Option<Receiver<DownloadEvent>>,
    /// Ergebnis des letzten Downloads für die Anzeige.
    download_status: Option<Result<String, String>>,
    /// Fenster mit der Spaltenzuordnung ist geöffnet.
    show_column_config: bool,
    /// Fenster mit den Rollen-Knöpfen ist geöffnet.
    show_role_config: bool,
    /// Meldung des Threads, der die Vorlage in Word öffnet.
    word_receiver: Option<Receiver<String>>,
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

    /// Prüft die OTP-Eingaben; solange der Schalter aus ist, werden keine Einmalpasswörter
    /// abgerufen. Ob alte OTPs gelöscht werden (`clear`), wird erst beim Start abgefragt.
    fn validate_otp(&self) -> Result<Option<OtpConfig>, String> {
        if !self.otp_enabled {
            return Ok(None);
        }
        let site_url = self.settings.otp_site_url.trim();
        if site_url.is_empty() {
            return Err("Bitte die WordPress-Adresse für die Einmalpasswörter eingeben.".to_string());
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
            clear: false,
        }))
    }

    /// Fragt vor einem Download mit Einmalpasswörtern, ob die alten OTPs auf dem Server
    /// gelöscht und neu erzeugt werden sollen. Liefert `None`, wenn abgebrochen wurde.
    /// Nutzt die Windows-Meldungsbox direkt, weil `rfd` keinen Standardknopf setzen kann;
    /// vorbelegt ist „Nein“, damit Enter nie versehentlich alle OTPs löscht.
    fn ask_clear_otp() -> Option<bool> {
        use windows::core::HSTRING;
        use windows::Win32::UI::WindowsAndMessaging::{
            MessageBoxW, IDNO, IDYES, MB_DEFBUTTON2, MB_ICONWARNING, MB_YESNOCANCEL,
        };
        let text = HSTRING::from(
            "Sollen die bestehenden Einmalpasswörter auf dem Server gelöscht und neu erzeugt \
             werden (clearotp=1)?\n\n\
             Ja: alte OTPs werden gelöscht und ungültig, alle Mitglieder erhalten neue.\n\
             Nein: alte OTPs bleiben gültig, es werden zusätzlich neue erzeugt.\n\
             Abbrechen: kein Download.",
        );
        let title = HSTRING::from("Einmalpasswörter");
        // SAFETY: MessageBoxW erhält gültige, nullterminierte Strings und kein Fensterhandle.
        let answer = unsafe {
            MessageBoxW(None, &text, &title, MB_YESNOCANCEL | MB_DEFBUTTON2 | MB_ICONWARNING)
        };
        match answer {
            IDYES => Some(true),
            IDNO => Some(false),
            _ => None,
        }
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

    /// Zeichnet das Fenster zum Bearbeiten der Rollen-Knöpfe (Beschriftung ↔ `role_list`).
    fn role_config_window(&mut self, ctx: &egui::Context, enabled: bool) {
        let mut open = self.show_role_config;
        egui::Window::new("Rollen-Knöpfe für den Admidio-Download")
            .open(&mut open)
            .default_width(640.0)
            .resizable(true)
            .collapsible(false)
            .show(ctx, |ui| {
                ui.label(
                    "Jede Zeile ergibt einen Knopf. Beim Klick wird in der Export-URL der \
                     Parameter role_list durch den angegebenen Wert ersetzt (bzw. ergänzt); der \
                     Download wird erst mit „CSV aus Admidio laden“ gestartet. Die übrigen \
                     Parameter wie list_uuid, date_from und date_to \
                     bleiben unverändert. In Admidio 5 ist role_list die UUID der Rolle; mehrere \
                     Rollen werden durch Komma getrennt.",
                );
                ui.add_space(6.0);

                let mut remove = None;
                egui::ScrollArea::vertical().max_height(360.0).show(ui, |ui| {
                    egui::Grid::new("role_buttons")
                        .num_columns(3)
                        .spacing([12.0, 6.0])
                        .striped(true)
                        .show(ui, |ui| {
                            ui.strong("Beschriftung (Rollenname)");
                            ui.strong("role_list (Rollen-UUID)");
                            ui.label("");
                            ui.end_row();
                            for (index, button) in self.settings.admidio_roles.iter_mut().enumerate() {
                                ui.add_enabled(
                                    enabled,
                                    egui::TextEdit::singleline(&mut button.name).desired_width(200.0),
                                );
                                ui.add_enabled(
                                    enabled,
                                    egui::TextEdit::singleline(&mut button.role_list).desired_width(320.0),
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
                    self.settings.admidio_roles.remove(index);
                }

                ui.add_space(6.0);
                ui.horizontal(|ui| {
                    if ui.add_enabled(enabled, egui::Button::new("Zeile hinzufügen")).clicked() {
                        self.settings.admidio_roles.push(RoleButton::default());
                    }
                    if ui
                        .add_enabled(
                            enabled && !self.settings.admidio_roles.is_empty(),
                            egui::Button::new("Alle entfernen"),
                        )
                        .clicked()
                    {
                        self.settings.admidio_roles.clear();
                    }
                });

                ui.add_space(4.0);
                match admidio::role_buttons(&self.settings.admidio_roles) {
                    Ok(buttons) => {
                        ui.label(format!("{} Rollen-Knöpfe werden angezeigt.", buttons.len()));
                    }
                    Err(error) => {
                        ui.colored_label(egui::Color32::RED, error.to_string());
                    }
                }
            });
        self.show_role_config = open;
    }

    /// Zeichnet die konfigurierten Rollen-Knöpfe; ein Klick trägt das `role_list` des Knopfs
    /// in die Export-URL ein. Der Download wird davon nicht gestartet.
    fn role_buttons_row(&mut self, ui: &mut egui::Ui, enabled: bool) {
        let buttons = match admidio::role_buttons(&self.settings.admidio_roles) {
            Ok(buttons) => buttons,
            Err(error) => {
                ui.colored_label(egui::Color32::RED, error.to_string());
                return;
            }
        };
        let has_url = !admidio::parse_urls(&self.settings.admidio_urls).is_empty();
        let current = admidio::query_param(&self.settings.admidio_urls, "role_list");
        let mut clicked = None;
        ui.horizontal_wrapped(|ui| {
            for button in &buttons {
                let active = current.as_deref() == Some(button.role_list.as_str());
                let response = Self::choice_button(ui, enabled && has_url, active, &button.name)
                    .on_hover_text(format!("Setzt role_list={} in der Export-URL", button.role_list));
                if response.clicked() {
                    clicked = Some(button.clone());
                }
            }
            if ui.button("Bearbeiten …").clicked() {
                self.show_role_config = true;
            }
            if buttons.is_empty() {
                ui.label("keine Rollen-Knöpfe konfiguriert");
            }
        });
        if let Some(button) = clicked {
            match admidio::replace_role_list(&self.settings.admidio_urls, &button.role_list) {
                Ok(urls) => {
                    self.settings.admidio_urls = urls;
                    self.download_status = None;
                    self.log.push(format!(
                        "Rolle '{}' gewählt, role_list={} in die Export-URL eingetragen",
                        button.name, button.role_list
                    ));
                }
                Err(error) => self.download_status = Some(Err(error.to_string())),
            }
        }
    }

    /// Fragt bei leerer Datenquelle nach einer Zieldatei und startet den Download.
    fn start_download(&mut self, mut config: AdmidioConfig) {
        if let Some(otp) = &mut config.otp {
            match Self::ask_clear_otp() {
                Some(clear) => otp.clear = clear,
                None => {
                    self.log.push("Download abgebrochen (Abfrage zu den Einmalpasswörtern).".to_string());
                    return;
                }
            }
        }
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
        let enabled = editable && !self.is_downloading();
        let username_missing = self.settings.admidio_username.trim().is_empty();
        let password_missing = self.admidio_password.is_empty();
        let urls = admidio::parse_urls(&self.settings.admidio_urls);
        let url_missing = urls.is_empty()
            || urls
                .iter()
                .any(|url| !(url.starts_with("https://") || url.starts_with("http://")) || admidio::login_url(url).is_err());
        let otp_url = self.settings.otp_site_url.trim();
        let otp_url_missing =
            self.otp_enabled && !(otp_url.starts_with("https://") || otp_url.starts_with("http://"));
        let otp_password_missing = self.otp_enabled && self.settings.otp_password.trim().is_empty();

        ui.group(|ui| {
            ui.set_width(ui.available_width());
            ui.strong("Admidio");
            ui.label("Anmeldung und Rollenliste, aus der die Datenquelle geladen wird.");
            ui.add_space(4.0);
            egui::Grid::new("admidio")
                .num_columns(2)
                .spacing([12.0, 8.0])
                .show(ui, |ui| {
                    ui.label("Benutzername");
                    let fill = Self::field_fill(ui, username_missing);
                    ui.add_enabled(
                        enabled,
                        egui::TextEdit::singleline(&mut self.settings.admidio_username)
                            .desired_width(f32::INFINITY)
                            .background_color(fill),
                    );
                    ui.end_row();

                    ui.label("Passwort");
                    let fill = Self::field_fill(ui, password_missing);
                    ui.add_enabled(
                        enabled,
                        egui::TextEdit::singleline(&mut self.admidio_password)
                            .password(true)
                            .desired_width(f32::INFINITY)
                            .background_color(fill)
                            .hint_text("wird nicht gespeichert"),
                    );
                    ui.end_row();

                    ui.label("Export-URL");
                    let fill = Self::field_fill(ui, url_missing);
                    ui.add_enabled(
                        enabled,
                        egui::TextEdit::multiline(&mut self.settings.admidio_urls)
                            .desired_rows(3)
                            .desired_width(f32::INFINITY)
                            .background_color(fill)
                            .hint_text(
                                "https://.../modules/groups-roles/lists_show.php?list_uuid=...&role_list=...&mode=csv",
                            ),
                    );
                    ui.end_row();

                    ui.label("Rolle wählen");
                    self.role_buttons_row(ui, enabled);
                    ui.end_row();

                    ui.label("Zeitraum");
                    self.date_buttons_row(ui, enabled);
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
                });
        });

        ui.add_space(4.0);
        ui.group(|ui| {
            ui.set_width(ui.available_width());
            ui.strong("WordPress (Einmalpasswörter)");
            ui.label(
                "Optional: holt die Einmalpasswörter von der WordPress-Seite und ergänzt die \
                 Spalten otp1 und otp2 (Zuordnung über username bzw. Mitgliedsnummer).",
            );
            ui.add_space(4.0);
            ui.add_enabled(
                enabled,
                egui::Checkbox::new(
                    &mut self.otp_enabled,
                    "Einmalpasswörter beim Download abrufen (wird nicht gespeichert)",
                ),
            );
            let otp_fields = enabled && self.otp_enabled;
            egui::Grid::new("wordpress")
                .num_columns(2)
                .spacing([12.0, 8.0])
                .show(ui, |ui| {
                    ui.label("WordPress-Adresse");
                    let fill = Self::field_fill(ui, otp_url_missing);
                    ui.add_enabled(
                        otp_fields,
                        egui::TextEdit::singleline(&mut self.settings.otp_site_url)
                            .desired_width(f32::INFINITY)
                            .background_color(fill)
                            .hint_text("https://<wordpress-seite>/"),
                    );
                    ui.end_row();

                    ui.label("OTP-Passwort");
                    let fill = Self::field_fill(ui, otp_password_missing);
                    ui.add_enabled(
                        otp_fields,
                        egui::TextEdit::singleline(&mut self.settings.otp_password)
                            .password(true)
                            .desired_width(f32::INFINITY)
                            .background_color(fill),
                    );
                    ui.end_row();
                });
            if self.otp_enabled {
                ui.label(
                    "Vor dem Download wird gefragt, ob die alten Einmalpasswörter auf dem Server \
                     gelöscht und neu erzeugt werden sollen (clearotp).",
                );
            }
        });

        ui.add_space(4.0);
        let validation = self.validate_download();
        ui.horizontal(|ui| {
            let can_start = enabled && validation.is_ok();
            if Self::action_button(ui, can_start, "CSV aus Admidio laden").clicked() {
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
        let flyer = self.settings.flyer_path.trim();
        let flyer = if flyer.is_empty() {
            None
        } else {
            let path = PathBuf::from(flyer);
            if !path.is_file() {
                return Err("Bitte eine vorhandene Flyer-PDF auswählen.".to_string());
            }
            Some(path)
        };
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
            flyer,
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
                Progress::Aborted(duration) => {
                    self.finished = Some(duration);
                    self.receiver = None;
                    self.log.push(format!(
                        "Abbruch nach {} ({} PDFs, {} Fehler)",
                        merge::format_duration(duration),
                        self.done,
                        self.failed
                    ));
                    return;
                }
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

    /// Zeichnet die Zeile „Zeitraum“: zeigt `date_from`/`date_to` der Export-URL und bietet den
    /// Knopf „Heute“, der beide auf das aktuelle Datum setzt. Grün, wenn die URL schon heute zeigt.
    fn date_buttons_row(&mut self, ui: &mut egui::Ui, enabled: bool) {
        let today = chrono::Local::now().format("%Y-%m-%d").to_string();
        let has_url = !admidio::parse_urls(&self.settings.admidio_urls).is_empty();
        let date_from = admidio::query_param(&self.settings.admidio_urls, "date_from");
        let date_to = admidio::query_param(&self.settings.admidio_urls, "date_to");
        let active = date_from.as_deref() == Some(today.as_str()) && date_to.as_deref() == Some(today.as_str());
        let mut clicked = false;
        ui.horizontal_wrapped(|ui| {
            clicked = Self::choice_button(ui, enabled && has_url, active, "Heute")
                .on_hover_text(format!("Setzt date_from={today} und date_to={today} in der Export-URL"))
                .clicked();
            match (&date_from, &date_to) {
                (Some(from), Some(to)) => {
                    ui.label(format!("in der URL: {from} bis {to}"));
                }
                _ if has_url => {
                    ui.label("in der URL nicht gesetzt, Admidio nimmt dann das heutige Datum");
                }
                _ => {}
            }
        });
        if clicked {
            let params = [("date_from", today.as_str()), ("date_to", today.as_str())];
            match admidio::replace_query_params(&self.settings.admidio_urls, &params) {
                Ok(urls) => {
                    self.settings.admidio_urls = urls;
                    self.download_status = None;
                    self.log.push(format!("Zeitraum auf {today} gesetzt (date_from und date_to in der Export-URL)"));
                }
                Err(error) => self.download_status = Some(Err(error.to_string())),
            }
        }
    }

    /// Zeichnet einen Auswahl-Knopf in normaler Größe. `active` hebt ihn grün hervor, weil
    /// sein Wert gerade in der Export-URL steht.
    fn choice_button(ui: &mut egui::Ui, enabled: bool, active: bool, text: &str) -> egui::Response {
        let button = if active {
            egui::Button::new(egui::RichText::new(text).strong().color(egui::Color32::WHITE))
                .fill(egui::Color32::from_rgb(46, 139, 69))
        } else {
            egui::Button::new(text)
        };
        ui.add_enabled(enabled, button)
    }

    /// Öffnet die Vorlage in Word und hängt die Datenquelle an, falls sie vorhanden und
    /// lesbar ist. Läuft in einem Thread, weil der Word-Start die Oberfläche sonst blockiert.
    fn open_in_word(&mut self, template: PathBuf) {
        let data_source = PathBuf::from(self.settings.data_source.trim());
        let data_source = (data_source.is_file() && self.data_error.is_none()).then_some(data_source);
        match &data_source {
            Some(path) => self.log.push(format!(
                "Word wird mit Vorlage {} und Datenquelle {} geöffnet ...",
                template.display(),
                path.display()
            )),
            None => self.log.push(format!(
                "Word wird mit Vorlage {} geöffnet; keine gültige Datenquelle, sie wird nicht angehängt.",
                template.display()
            )),
        }
        let (tx, rx) = std::sync::mpsc::channel();
        self.word_receiver = Some(rx);
        std::thread::spawn(move || {
            let message = match merge::open_in_word(&template, data_source.as_deref()) {
                Ok(()) => "Word geöffnet.".to_string(),
                Err(error) => format!("Word konnte nicht geöffnet werden: {error:#}"),
            };
            let _ = tx.send(message);
        });
    }

    /// Holt die Meldung des Word-Threads ab.
    fn poll_word(&mut self) {
        let Some(receiver) = &self.word_receiver else {
            return;
        };
        match receiver.try_recv() {
            Ok(message) => {
                self.log.push(message);
                self.word_receiver = None;
            }
            Err(std::sync::mpsc::TryRecvError::Disconnected) => self.word_receiver = None,
            Err(std::sync::mpsc::TryRecvError::Empty) => {}
        }
    }

    /// Öffnet eine Datei mit dem in Windows verknüpften Programm (für `.docx` also Word),
    /// ohne ein Konsolenfenster zu zeigen.
    fn open_with_default_app(path: &Path) -> std::io::Result<()> {
        use std::os::windows::process::CommandExt;
        const CREATE_NO_WINDOW: u32 = 0x0800_0000;
        std::process::Command::new("cmd")
            .args(["/C", "start", ""])
            .arg(path)
            .creation_flags(CREATE_NO_WINDOW)
            .spawn()
            .map(|_| ())
    }

    /// Zeichnet einen Knopf, der eine Aktion auslöst: größere Schrift und mehr Innenabstand
    /// als die übrigen Knöpfe. Grün hinterlegt, wenn die Aktion möglich ist, sonst grau.
    fn action_button(ui: &mut egui::Ui, enabled: bool, text: &str) -> egui::Response {
        let (fill, text_color) = if enabled {
            (egui::Color32::from_rgb(46, 139, 69), egui::Color32::WHITE)
        } else {
            (egui::Color32::from_gray(170), egui::Color32::from_gray(60))
        };
        ui.scope(|ui| {
            ui.spacing_mut().button_padding = egui::vec2(14.0, 7.0);
            let label = egui::RichText::new(text).size(16.0).strong().color(text_color);
            ui.add_enabled(
                enabled,
                egui::Button::new(label)
                    .fill(fill)
                    .min_size(egui::vec2(0.0, 34.0)),
            )
        })
        .inner
    }

    /// Hintergrundfarbe eines Pflichtfelds: leicht rosa, wenn die Eingabe fehlt oder ungültig
    /// ist, sonst die normale Feldfarbe. Im dunklen Design ein gedecktes Rot.
    fn field_fill(ui: &egui::Ui, missing: bool) -> egui::Color32 {
        if !missing {
            ui.visuals().extreme_bg_color
        } else if ui.visuals().dark_mode {
            egui::Color32::from_rgb(96, 40, 52)
        } else {
            egui::Color32::from_rgb(255, 226, 232)
        }
    }

    /// Zeichnet eine Zeile mit Beschriftung, Textfeld und Auswahl-Knopf; liefert `true`,
    /// wenn sich der Wert geändert hat. `missing` hinterlegt das Feld rosa.
    fn path_row(
        ui: &mut egui::Ui,
        label: &str,
        value: &mut String,
        enabled: bool,
        missing: bool,
        pick: impl FnOnce() -> Option<PathBuf>,
    ) -> bool {
        let mut changed = false;
        let fill = Self::field_fill(ui, missing);
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
                    .add(
                        egui::TextEdit::singleline(value)
                            .desired_width(f32::INFINITY)
                            .background_color(fill),
                    )
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
        self.poll_word();
        if self.is_running() || self.is_downloading() || self.word_receiver.is_some() {
            ctx.request_repaint_after(Duration::from_millis(200));
        }
        let editable = !self.is_running();

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Word Serienbrief als PDF");
            ui.add_space(8.0);

            let mut data_changed = false;
            let mut template_changed = false;
            let template_missing = !Path::new(self.settings.template.trim()).is_file();
            let data_missing = !Path::new(self.settings.data_source.trim()).is_file() || self.data_error.is_some();
            let output_missing = self.settings.output_dir.trim().is_empty();
            let flyer = self.settings.flyer_path.trim();
            let flyer_missing = !flyer.is_empty() && !Path::new(flyer).is_file();
            let pattern = self.settings.file_pattern.trim();
            let pattern_missing = pattern.is_empty()
                || merge::placeholders(pattern).is_empty()
                || !self.unknown_placeholders().is_empty();

            egui::Grid::new("settings")
                .num_columns(2)
                .spacing([12.0, 8.0])
                .show(ui, |ui| {
                    template_changed = Self::path_row(
                        ui,
                        "Serienbrief (Word)",
                        &mut self.settings.template,
                        editable,
                        template_missing,
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
                        data_missing,
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
                        output_missing,
                        || rfd::FileDialog::new().pick_folder(),
                    );
                    Self::path_row(
                        ui,
                        "Flyer anhängen (PDF, optional)",
                        &mut self.settings.flyer_path,
                        editable,
                        flyer_missing,
                        || {
                            rfd::FileDialog::new()
                                .add_filter("PDF-Dateien", &["pdf"])
                                .add_filter("Alle Dateien", &["*"])
                                .pick_file()
                        },
                    );

                    ui.label("Dateiname");
                    let pattern_fill = Self::field_fill(ui, pattern_missing);
                    ui.add_enabled_ui(editable, |ui| {
                        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                            if ui.button("Standard").clicked() {
                                self.settings.file_pattern = DEFAULT_PATTERN.to_string();
                            }
                            ui.add(
                                egui::TextEdit::singleline(&mut self.settings.file_pattern)
                                    .desired_width(f32::INFINITY)
                                    .background_color(pattern_fill),
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
                if Self::action_button(ui, can_start, "PDFs erzeugen").clicked() {
                    if let Ok(config) = &validation {
                        self.start(config.clone());
                    }
                }
                let template = PathBuf::from(self.settings.template.trim());
                let can_open = editable && template.is_file() && self.word_receiver.is_none();
                if Self::action_button(ui, can_open, "Word öffnen")
                    .on_hover_text("Öffnet die Serienbrief-Vorlage in Word und hängt die Datenquelle an")
                    .clicked()
                {
                    self.open_in_word(template);
                }
                // Der Zielordner darf auch während des Laufs geöffnet werden, um die
                // entstehenden PDFs zu sehen.
                let output_dir = PathBuf::from(self.settings.output_dir.trim());
                if Self::action_button(ui, output_dir.is_dir(), "Zielordner öffnen")
                    .on_hover_text("Öffnet den Zielordner im Explorer")
                    .clicked()
                {
                    match Self::open_with_default_app(&output_dir) {
                        Ok(()) => self.log.push(format!("Zielordner geöffnet: {}", output_dir.display())),
                        Err(error) => self.log.push(format!("Zielordner konnte nicht geöffnet werden: {error}")),
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
        self.role_config_window(ctx, editable && !self.is_downloading());
    }
}

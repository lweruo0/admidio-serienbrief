//! Late-Binding-Zugriff auf Microsoft Word über die COM-Schnittstelle IDispatch.
//!
//! Die Klasse `ComObject` kapselt ein beliebiges Automation-Objekt (Word.Application,
//! Document, MailMerge, ...) und erlaubt Property-Zugriffe und Methodenaufrufe über den
//! Namen, so wie es `win32com.client.Dispatch` in Python macht.

use anyhow::{anyhow, Context, Result};
use windows::core::{BSTR, GUID, HSTRING, PCWSTR, VARIANT};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize, IDispatch,
    CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPATCH_METHOD,
    DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
};

/// DISPID des Wertes bei Property-Put-Aufrufen (siehe oaidl.h).
const DISPID_PROPERTYPUT: i32 = -3;

/// Locale für IDispatch-Aufrufe (LOCALE_USER_DEFAULT).
const LOCALE_USER_DEFAULT: u32 = 0x0400;

/// Initialisiert COM für den aktuellen Thread und gibt es beim Verlassen des Gültigkeitsbereichs
/// wieder frei. Jeder Arbeits-Thread benötigt eine eigene Instanz.
pub struct ComApartment;

impl ComApartment {
    /// Initialisiert ein Single-Threaded-Apartment für den aktuellen Thread.
    pub fn new() -> Result<Self> {
        // SAFETY: CoInitializeEx darf pro Thread aufgerufen werden; der zugehörige
        // CoUninitialize-Aufruf erfolgt im Drop.
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() }
            .context("COM konnte nicht initialisiert werden")?;
        Ok(Self)
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        // SAFETY: Gegenstück zu CoInitializeEx im Konstruktor.
        unsafe { CoUninitialize() };
    }
}

/// Dünner Wrapper um ein IDispatch-Objekt mit Zugriff über Elementnamen.
pub struct ComObject {
    dispatch: IDispatch,
}

impl ComObject {
    /// Erzeugt ein neues Automation-Objekt anhand seiner ProgID, z. B. `Word.Application`.
    pub fn create(prog_id: &str) -> Result<Self> {
        let prog_id_wide = HSTRING::from(prog_id);
        // SAFETY: Beide Funktionen erhalten gültige Zeiger; COM ist über ComApartment initialisiert.
        let dispatch: IDispatch = unsafe {
            let clsid = CLSIDFromProgID(&prog_id_wide)
                .with_context(|| format!("ProgID '{prog_id}' ist nicht registriert"))?;
            CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)
                .with_context(|| format!("'{prog_id}' konnte nicht gestartet werden"))?
        };
        Ok(Self { dispatch })
    }

    /// Liest eine Eigenschaft.
    pub fn get(&self, name: &str) -> Result<VARIANT> {
        self.invoke(name, DISPATCH_PROPERTYGET, Vec::new())
    }

    /// Liest eine Eigenschaft, die ein weiteres Automation-Objekt liefert.
    pub fn get_object(&self, name: &str) -> Result<ComObject> {
        Self::from_variant(self.get(name)?, name)
    }

    /// Liest eine Eigenschaft als Zeichenkette.
    pub fn get_string(&self, name: &str) -> Result<String> {
        let value = self.get(name)?;
        BSTR::try_from(&value)
            .map(|text| text.to_string())
            .with_context(|| format!("Eigenschaft '{name}' ist keine Zeichenkette"))
    }

    /// Setzt eine Eigenschaft.
    pub fn put(&self, name: &str, value: impl Into<VARIANT>) -> Result<()> {
        self.invoke(name, DISPATCH_PROPERTYPUT, vec![value.into()])?;
        Ok(())
    }

    /// Ruft eine Methode (oder eine Eigenschaft mit Parametern) auf.
    pub fn call(&self, name: &str, args: Vec<VARIANT>) -> Result<VARIANT> {
        self.invoke(name, DISPATCH_METHOD | DISPATCH_PROPERTYGET, args)
    }

    /// Ruft eine Methode auf, deren Rückgabe ein weiteres Automation-Objekt ist.
    pub fn call_object(&self, name: &str, args: Vec<VARIANT>) -> Result<ComObject> {
        Self::from_variant(self.call(name, args)?, name)
    }

    /// Wandelt einen VARIANT mit IDispatch-Inhalt in ein `ComObject`.
    fn from_variant(value: VARIANT, name: &str) -> Result<ComObject> {
        let dispatch = IDispatch::try_from(&value)
            .with_context(|| format!("'{name}' lieferte kein Automation-Objekt"))?;
        Ok(ComObject { dispatch })
    }

    /// Ermittelt die DISPID zu einem Elementnamen.
    fn dispid(&self, name: &str) -> Result<i32> {
        let name_wide = HSTRING::from(name);
        let names = [PCWSTR(name_wide.as_ptr())];
        let mut dispid = 0i32;
        // SAFETY: Alle Zeiger verweisen auf gültige lokale Daten.
        unsafe {
            self.dispatch.GetIDsOfNames(
                &GUID::zeroed(),
                names.as_ptr(),
                1,
                LOCALE_USER_DEFAULT,
                &mut dispid,
            )
        }
        .with_context(|| format!("Unbekanntes COM-Element '{name}'"))?;
        Ok(dispid)
    }

    /// Führt den eigentlichen IDispatch::Invoke aus.
    fn invoke(&self, name: &str, flags: DISPATCH_FLAGS, mut args: Vec<VARIANT>) -> Result<VARIANT> {
        let dispid = self.dispid(name)?;

        // IDispatch erwartet die Argumente in umgekehrter Reihenfolge.
        args.reverse();

        let is_put = flags == DISPATCH_PROPERTYPUT;
        let mut named_arg = DISPID_PROPERTYPUT;
        let params = DISPPARAMS {
            rgvarg: args.as_mut_ptr(),
            rgdispidNamedArgs: if is_put { &mut named_arg } else { std::ptr::null_mut() },
            cArgs: args.len() as u32,
            cNamedArgs: if is_put { 1 } else { 0 },
        };

        let mut result = VARIANT::default();
        let mut exception = EXCEPINFO::default();

        // SAFETY: params, result und exception leben bis nach dem Aufruf.
        let outcome = unsafe {
            self.dispatch.Invoke(
                dispid,
                &GUID::zeroed(),
                LOCALE_USER_DEFAULT,
                flags,
                &params,
                Some(&mut result),
                Some(&mut exception),
                None,
            )
        };

        match outcome {
            Ok(()) => Ok(result),
            Err(error) => {
                let description = exception.bstrDescription.to_string();
                let detail = if description.trim().is_empty() {
                    error.to_string()
                } else {
                    description.trim().to_string()
                };
                Err(anyhow!("COM-Aufruf '{name}' fehlgeschlagen: {detail}"))
            }
        }
    }
}

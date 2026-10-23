import os
import subprocess

# Pfad zum Ordner mit den PDFs
pdf_ordner = r"P:\BFV-Schriftführer\Rundschreiben\2026\oktober"  # <-- Ändere das zu deinem Pfad
ausgabe_ordner = os.path.join(pdf_ordner, "flyer")

# Ausgabeordner erstellen, falls nicht vorhanden
os.makedirs(ausgabe_ordner, exist_ok=True)


flyerpfad = r"P:\BFV-Schriftführer\Rundschreiben\2026\flyer_dezember.pdf"
# Alle PDF-Dateien im Verzeichnis durchgehen
for dateiname in os.listdir(pdf_ordner):
    if dateiname.lower().endswith(".pdf"):
        eingabepfad = os.path.join(pdf_ordner, dateiname)
        ausgabepfad = os.path.join(ausgabe_ordner, f"{os.path.splitext(dateiname)[0]}.pdf")

        # Mit pdftk nur die erste Seite extrahieren
        subprocess.run(["pdftk", f"A={eingabepfad}" , f"B={flyerpfad}" , "cat", "A1-2", "B1", "output", ausgabepfad])

print("Fertig! Alle PDFs wurden auf die erste Seite gekürzt.")

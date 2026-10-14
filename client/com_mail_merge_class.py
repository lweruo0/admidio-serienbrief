# -*- coding: utf-8 -*-

import os
import datetime
import time
import threading
import shutil
import subprocess
import win32com.client  # pylint: disable=import-error
import pythoncom  # pylint: disable=import-error


def thread_function(index, doc_template_name, data_source_name, chunk):
    job = MailMergeJob(index, doc_template_name, data_source_name, chunk)
    job.run()
    job.cleanup()


class MailMergeJob:
    def __init__(self, index, doc_template_name, data_source_name, chunk):

        temp_dir = os.environ.get("TEMP", "/tmp")
        doc_copy = os.path.join(temp_dir, f"rundschreiben_{index}.docx")
        data_copy = os.path.join(temp_dir, f"data_{index}.csv")

        shutil.copyfile(doc_template_name, doc_copy)
        shutil.copyfile(data_source_name, data_copy)

        self.doc_template_name = os.path.abspath(doc_copy)
        self.data_source_name = os.path.abspath(data_copy)
        self.chunk = chunk
        self.index = index

        zielpfad = os.path.dirname(doc_template_name)
        zielpfad = os.path.join(zielpfad, datetime.datetime.now().strftime("%Y-%m-%d-Rundschreiben"))
        if not os.path.exists(zielpfad):
            os.mkdir(zielpfad)

        self.zielpfad = zielpfad

    def run(self):

        app = win32com.client.Dispatch("Word.Application", pythoncom.CoInitialize())
        doc_template = app.Documents.Open(self.doc_template_name)
        mm = doc_template.MailMerge

        # attach data source to template
        mm.OpenDataSource(self.data_source_name)

        # send the merge result to a new document
        mm.Destination = 0  # win32com.client.constants.wdSendToNewDocument (0) wdSendToPrinter (1)

        st = time.time()

        anzahl_datensaetze = len(self.chunk)
        a = 1
        for i in self.chunk:
            mm.DataSource.FirstRecord = i
            mm.DataSource.LastRecord = i
            mm.DataSource.ActiveRecord = mm.DataSource.FirstRecord

            username = mm.DataSource.DataFields("username")
            Mitgliedsnummer = mm.DataSource.DataFields("Mitgliedsnummer")

            doc_final_name = os.path.join(self.zielpfad, f"rundschreiben-{Mitgliedsnummer}-{username}.pdf")
            dauer_sek = time.time() - st
            f = (anzahl_datensaetze - a) / a
            a += 1

            dauer_str = datetime.datetime.fromtimestamp(23 * 60 * 60 + dauer_sek).strftime("%M:%S")
            noch_str = datetime.datetime.fromtimestamp(23 * 60 * 60 + dauer_sek * f).strftime("%M:%S")

            print("%s verbleibend: %s  %s" % (dauer_str, noch_str, doc_final_name))

            mm.Execute()
            doc_final = app.Documents[0]
            doc_final.SaveAs(doc_final_name, FileFormat=17)  # wdFormatPDF = 17
            doc_final.Close(SaveChanges=False)

        doc_template.Close(SaveChanges=False)
        app.Quit()

    def cleanup(self):
        if os.path.exists(self.doc_template_name):
            os.remove(self.doc_template_name)
        if os.path.exists(self.data_source_name):
            os.remove(self.data_source_name)


if __name__ == "__main__":

    subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], shell=True)

    doc_template_name = os.path.abspath(
        r"P:\BFV-Schriftführer\Rundschreiben\2025\Rundschreiben_oktober\2025_10_xx_Rundschreiben.docx"
    )
    data_source_name = os.path.abspath(r"C:\Users\Oliver\OneDrive\Dokumente\Meine Datenquellen\_export_alle.csv")

    st_gesammt = time.time()

    anzahl_threads = 8

    inhalt = [x for x in open(data_source_name).readlines() if x.strip() != ""]
    anzahl_datensaetze = len(inhalt) - 1
    # Teile indices in anzahl_threads möglichst gleich große Stücke auf
    indices = list(range(1, anzahl_datensaetze + 1))
    chunk_size = (len(indices) + anzahl_threads - 1) // anzahl_threads
    indices_chunks = [indices[i * chunk_size : (i + 1) * chunk_size] for i in range(anzahl_threads)]

    threads = list()
    for index, chunk in enumerate(indices_chunks):
        x = threading.Thread(target=thread_function, args=(index, doc_template_name, data_source_name, chunk))
        threads.append(x)
        x.start()

    for thread in threads:
        thread.join()
    dauer_sek = time.time() - st_gesammt
    print("Gesamtdauer: %s" % datetime.datetime.fromtimestamp(dauer_sek).strftime("%M:%S"))

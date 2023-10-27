"""
Leihschein.py: Hilfsskript zur einheitlichen Anlage von Leihscheinen.
Alle Formularfelder werden abgefragt und ausgefüllt.
Zusätzlich wird ein Outlook-Task mit Ablaufdatum angelegt und das generierte PDF
zur Überprüfung und anschließenden Druck automatisch geöffnet.
"""

__author__ = "Florian Stöckl"
___version__ = "1.1"
__maintainer__ = "Florian Stöckl"
__email__ = "florian.stoeckl@dhbw-karlsruhe.de"
__status__ = "Production"

import os
import sys
from datetime import datetime
from typing import Optional

import win32com.client
from pypdf import PdfReader, PdfWriter


class Artikel:
    """
    Klasse für entliehene Artikel
    """

    def __init__(self, p_pos: Optional[int] = None, menge: int = None, bezeichnung: str = None,
                 seriennummer: Optional[str] = None, inventar_nummer: Optional[str] = None):
        self.pos = p_pos
        self.menge = menge
        self.bezeichnung = bezeichnung
        self.seriennummer = seriennummer
        self.inventar_nummer = inventar_nummer


def generate_uniform_leihschein_filename(p_name, p_rueckgabedatum):
    """
    Hilfsmethoden zur Erstellung eines einheitlichen Dateinamens für Leihscheine
    :param p_name: Name der ausleihenden Person
    :param p_rueckgabedatum: Rückgabedatum aller Artikel
    :return: Einheitlicher Dateiname im Format <Heutiges Datum>-<Rückgabedatum>-Leihschein-<Name der ausleihenden Person>.pdf
    """
    current_date_YYYY_MM_DD = datetime.now().strftime('%Y-%m-%d')
    rueckgabedatum_datetime = datetime.strptime(p_rueckgabedatum, "%d.%m.%Y")
    rueckgabedatum_formatted = rueckgabedatum_datetime.strftime('%Y-%m-%d')

    return f'{current_date_YYYY_MM_DD}-{rueckgabedatum_formatted}-Leihschein-{p_name}.pdf'


def generate_leihschein_pdf(p_studiengang, p_name, p_kurs, p_email, p_ausgeliehene_artikel, p_rueckgabedatum, p_verwendungszweck,
                            p_ausgegeben_durch, p_leihdatum, p_leihschein_filename):
    """
    Generiert basierend auf einem Template und den eingegebenen Informationen einen ausgefüllten Leihschein
    :param p_studiengang: Name des Studiengangs der verleihenden Person
    :param p_name: Name der ausleihenden Person
    :param p_kurs: Kurs der ausleihenden Person
    :param p_email: E-Mail-Adresse der ausleihenden Person
    :param p_ausgeliehene_artikel: Ausgeliehene Artikel
    :param p_rueckgabedatum: Rückgabedatum aller Artikel
    :param p_verwendungszweck: Verwendungszweck
    :param p_ausgegeben_durch: Ausgebende Person
    :param p_leihdatum: Datum der Ausleihe (in der Regel heutiges Datum)
    :param p_leihschein_filename: Gewünschter Dateiname des Leihschein-PDFs
    :return: Keine Rückgabe
    """
    # Original PDF template file
    template_file = os.path.join(application_path, 'Ausleihe_leer.pdf')

    # Open the output PDF in append and read mode
    with open(p_leihschein_filename, 'wb') as output_pdf:
        # Open the template PDF in read mode
        with open(template_file, 'rb') as template_pdf:
            # Create a PDF reader and writer objects
            reader = PdfReader(template_pdf)
            writer = PdfWriter()

            writer.append(reader)

            writer.update_page_form_field_values(
                writer.pages[0],
                {
                    "Studiengang": p_studiengang,
                    "Name": p_name,
                    "Kurs": p_kurs,
                    "Email": p_email,
                    "Rückgabedatum": p_rueckgabedatum,
                    "Verwendungszweck": p_verwendungszweck,
                    "Ausgegeben durch": p_ausgegeben_durch,
                    "Datum": p_leihdatum
                }
            )
            ausgeliehene_artikel_line = 0
            for p_artikel in p_ausgeliehene_artikel:
                artikel_appendix = ''
                if ausgeliehene_artikel_line > 0:
                    artikel_appendix = '_' + str(ausgeliehene_artikel_line)

                writer.update_page_form_field_values(
                    writer.pages[0],
                    {
                        "Pos" + artikel_appendix: p_artikel.pos,
                        "Menge" + artikel_appendix: p_artikel.menge,
                        "Bezeichnung" + artikel_appendix: p_artikel.bezeichnung,
                        "Seriennummer" + artikel_appendix: p_artikel.seriennummer,
                        "Inventar-Nummer" + artikel_appendix: p_artikel.inventar_nummer
                    }
                )
                ausgeliehene_artikel_line += 1

            # Save the changes to the output PDF
            writer.write(output_pdf)

    print(f'Leihschein PDF erfolgreich erstellt: {p_leihschein_filename}')


def create_outlook_task_as_reminder(p_name, p_kurs, p_email, p_ausgeliehene_artikel, p_rueckgabedatum,
                                    p_verwendungszweck, p_leihdatum):
    """
    Hilfsmethode zur Erstellung eines Outlook Tasks mit Ablaufdatum=Rückgabedatum und allen relevanten Informationen zur Ausleihe
    :param p_name: Name der ausleihenden Person
    :param p_kurs: Kurs der ausleihenden Person
    :param p_email: E-Mail-Adresse der ausleihenden Person
    :param p_ausgeliehene_artikel: Ausgeliehene Artikel
    :param p_rueckgabedatum: Rückgabedatum aller Artikel
    :param p_verwendungszweck: Verwendungszweck
    :param p_leihdatum: Datum der Ausleihe (in der Regel heutiges Datum)
    :return:
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    tasks_folder = outlook.getDefaultFolder(13)
    task = tasks_folder.Items.Add()

    ausgeliehene_artikel_body = ''
    for artikel in p_ausgeliehene_artikel:
        ausgeliehene_artikel_body += artikel.bezeichnung + ', '

    rueckgabedatum_datetime = datetime.strptime(p_rueckgabedatum, "%d.%m.%Y")
    rueckgabedatum_formatted = rueckgabedatum_datetime.strftime('%m/%d/%Y')

    task.Subject = 'Rückgabe des Leihscheins ' + generate_uniform_leihschein_filename(p_name, p_rueckgabedatum)
    task.Body = f'Die Rückgabe von {p_name} ({p_email}) aus dem Kurs {p_kurs} ist fällig. Folgende Artikel wurden am {p_leihdatum} für {p_verwendungszweck} geliehen: \n{ausgeliehene_artikel_body}'
    task.DueDate = rueckgabedatum_formatted

    task.Save()


def open_pdf_file(p_leihschein_filename):
    """
    Öffnet PDF im Standardprogramm
    :param p_leihschein_filename:
    :return:
    """
    os.startfile(
        # print command als zweiter Parameter möglich, startet direkten Druck mit primärem Standarddrucker
        p_leihschein_filename)


if __name__ == "__main__":
    # Prüfe, wie das Python-Skript ausgeführt wird, um Applikationspfad korrekt zu setzen
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        application_path = sys._MEIPASS
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    # Eingabeaufforderungen für alle Formularfelder
    studiengang = input('Studiengang: ')
    name = input('Name: ')
    kurs = input('Kurs: ')
    email = input('Email: ')
    anzahl_ausgeliehene_artikel = int(input('Anzahl ausgeliehene Artikel: '))
    ausgeliehene_artikel = [Artikel() for i in range(anzahl_ausgeliehene_artikel)]
    for artikel in ausgeliehene_artikel:
        pos = input('Pos.: ')
        artikel.pos = int(pos) if pos else None
        artikel.menge = int(input('Menge: '))
        artikel.bezeichnung = input('Bezeichnung: ')
        artikel.seriennummer = input('Seriennummer: ')
        artikel.inventar_nummer = input('Inventar-Nummer: ')
    rueckgabedatum = input('Rückgabedatum: ')
    verwendungszweck = input('Verwendungszweck: ')
    ausgegeben_durch = input('Ausgegeben durch: ')
    leihdatum = datetime.now().strftime('%d.%m.%Y')

    leihschein_filename = generate_uniform_leihschein_filename(name, rueckgabedatum)

    # Generiere Leihschein PDF
    generate_leihschein_pdf(studiengang, name, kurs, email, ausgeliehene_artikel, rueckgabedatum, verwendungszweck, ausgegeben_durch,
                            leihdatum, leihschein_filename)
    create_outlook_task_as_reminder(name, kurs, email, ausgeliehene_artikel, rueckgabedatum, verwendungszweck,
                                    leihdatum)
    open_pdf_file(leihschein_filename)

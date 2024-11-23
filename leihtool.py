"""
Leihschein.py: Hilfsskript zur einheitlichen Anlage von Leihscheinen.
Alle Formularfelder werden abgefragt und ausgefüllt.
Zusätzlich wird ein Outlook-Task mit Ablaufdatum angelegt und das generierte PDF
zur Überprüfung und anschließenden Druck automatisch geöffnet.
"""

__author__ = "Florian Stöckl"
___version__ = "1.3"
__maintainer__ = "Florian Stöckl"
__email__ = "florian.stoeckl@dhbw-karlsruhe.de"
__status__ = "Production"

import os
import sys
from datetime import datetime
from typing import Optional
import re
import keyboard

from prompt_toolkit.document import Document
import win32com.client
from pypdf import PdfReader, PdfWriter
import questionary
from questionary import Validator, ValidationError


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

class NameValidator(Validator):
    def validate(self, document: Document) -> None:
        if len(document.text) == 0:
            raise ValidationError(message="Bitte geben Sie einen Namen ein.")
        
class AnzahlValidator(Validator):
    def validate(self, document: Document) -> None:
        if not document.text.isdigit() or int(document.text) < 1: 
            raise ValidationError(message="Bitte geben Sie eine Zahl größer 0 ein. Sie füllen gerade einen Leihschein aus, um mindestens einen Artikel zu verleihen.")

class NumberValidator(Validator):
    def validate(self, document: Document) -> None:
        if not document.text.isdigit(): 
            raise ValidationError(message="Bitte geben Sie eine Zahl ein.")
        
class EMailValidator(Validator):
    def validate(self, document: Document) -> None:
        if not re.fullmatch(r"[^@]+@[^@]+\.[^@]+", document.text): 
            raise ValidationError(message="Bitte geben Sie eine gültige E-Mail-Adresse ein.")
        
class DateValidator(Validator):
    def validate(self, document: Document) -> None:
        if re.fullmatch(r"[0-3][0-9]\.[0-1][0-9]\.[0-9]{4}", document.text):
            try:
               return_date = datetime.strptime(document.text, "%d.%m.%Y").date()
            except ValueError as exc:
                raise ValidationError(message="Bitte geben Sie ein gültiges Datum im Format DD.MM.YYYY ein.") from exc
            now = datetime.now().date()
            if return_date < now:
               raise ValidationError(message="Das Rückgabedatum kann nicht in der Vergangenheit liegen.")
        else:
            raise ValidationError(message="Bitte geben Sie ein gültiges Datum im Format DD.MM.YYYY ein.")

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
    :return: keine Rückgabe
    """
    os.startfile(
        # print command als zweiter Parameter möglich, startet direkten Druck mit primärem Standarddrucker
        p_leihschein_filename)


def send_email_to_lender(p_email, p_leihschein_filename):
    """
    Versendet über Outlook eine E-Mail an die ausleihende Person
    :param p_mail: E-Mail der ausleihenden Person
    :param p_leihschein_filename: Dateiname des generierten Leihschein-PDFs
    :return: keine Rückgabe
    """
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = p_email
    mail.Subject = 'Leihschein für entliehene Artikel der DHBW Karlsruhe'
    message = 'Guten Tag,<br>' \
                'Sie haben soeben Artikel der DHBW Karlsruhe entliehen.<br>' \
                'Dazu haben Sie einen Leihschein ausgefüllt und unterschrieben, den Sie im Anhang finden.<br>' \
                'Er enthält alle relevanten Informationen wie die ausgeliehenen Artikel und das Rückgabedatum.<br>' \
                '<br>' \
                'Bitte geben Sie die Artikel zum vereinbarten Rückgabetermin an die Person zurück, von der Sie die Artikel ausgeliehen haben.<br>' \
                'Vielen Dank im Voraus.<br>' \
                '<br>' \
                'Mit freundlichen Grüßen'
    # Add default Signature, if available
    mail.Display(False)
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    mail.HTMLbody = mail.HTMLbody[:index + 1] + message + mail.HTMLbody[index + 1:] 


    absolut_path_to_leihschein_pdf = os.path.join(execution_path, p_leihschein_filename)
    mail.Attachments.Add(absolut_path_to_leihschein_pdf)

    mail.Send()

def show_menu():
    menu_options = [
        "Standardwerte festlegen",
        "Einstellungen für externe Anwendungen (Obsidian)",
        "Programm beenden"
    ]
    choice = questionary.select(
        "Wählen Sie eine Option:",
        choices=menu_options
    ).ask()

    if choice == "Standardwerte festlegen":
        # Aktion für Option 1
        pass
    elif choice == "Einstellungen für externe Anwendungen (Obsidian)":
        # Aktion für Option 2
        pass
    elif choice == "Programm beenden":
        sys.exit(0)

# Funktion, die auf F1-Taste reagiert
def on_f1_press(event):
    sys.stdout.write('\033[2K\033[1G')  # Löscht die aktuelle Zeile im Terminal
    print("F1 wurde gedrückt. Öffne Menü...")
    show_menu()

# Überwachen der F1-Taste
keyboard.on_press_key("F1", on_f1_press)



if __name__ == "__main__":
    # Prüfe, wie das Python-Skript ausgeführt wird, um Applikationspfad korrekt zu setzen
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        application_path = sys._MEIPASS
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    execution_path = os.path.dirname(sys.executable)

    # Programmheader
    print(f"{'*' * 60}\n Sie verwenden das Leihscheintool in Version {___version__}\n" 
           + " Es hilft Ihnen beim Verleih von Hardware der DHBW.\n"
           + "\n"
           + " Sie können jederzeit mit F1 ins Menü wechseln,\n"
            + " um Einstellungen vorzunehmen.\n"
          + f"{'*' * 60}")

    # Eingabeaufforderungen für alle Formularfelder
    studiengang = questionary.autocomplete(
        'Studiengang:',
        choices=[
            'Informatik',
            'Maschinenbau',
            'Gesundheitswesen',
            'Mechatronik',
            'BWL',
            'Sozialwesen',
            'Angewandte Hebammenwissenschaft',
            'Soziale Arbeit',
            'Architektur',
            'Bauingenieurwesen',
            'Elektrotechnik',
            'Medizintechnik',
            'Wirtschaftsingenierurwesen',
            'Medien'
        ]).ask()
    name = questionary.text('Name:', validate=NameValidator).ask()
    kurs = questionary.text('Kurs:').ask()
    email = questionary.text('Email:', validate=EMailValidator).ask()
    anzahl_ausgeliehene_artikel = int(questionary.text('Anzahl ausgeliehene Artikel:', validate=AnzahlValidator).ask())
    ausgeliehene_artikel = [Artikel() for i in range(anzahl_ausgeliehene_artikel)]
    for artikel in ausgeliehene_artikel:
        questionary.print(f"Eingabe des {ausgeliehene_artikel.index(artikel) + 1}. Artikels:", style='bold')
        artikel.pos = int(questionary.text('Pos.:', validate=NumberValidator).ask())
        artikel.menge = int(questionary.text('Menge:', validate=NumberValidator).ask())
        artikel.bezeichnung = questionary.text('Bezeichnung:').ask()
        artikel.seriennummer = questionary.text('Seriennummer:').ask()
        artikel.inventar_nummer = questionary.text('Inventar-Nummer:').ask()
    rueckgabedatum = questionary.text('Rückgabedatum:', validate=DateValidator).ask()
    verwendungszweck = questionary.text('Verwendungszweck:').ask()
    ausgegeben_durch = questionary.text('Ausgegeben durch:').ask()
    leihdatum = datetime.now().strftime('%d.%m.%Y')

    leihschein_filename = generate_uniform_leihschein_filename(name, rueckgabedatum)

    # Generiere Leihschein PDF
    generate_leihschein_pdf(studiengang, name, kurs, email, ausgeliehene_artikel, rueckgabedatum, verwendungszweck, ausgegeben_durch,
                            leihdatum, leihschein_filename)
    # create_outlook_task_as_reminder(name, kurs, email, ausgeliehene_artikel, rueckgabedatum, verwendungszweck,
    #                                 leihdatum)
    # send_email_to_lender(email, leihschein_filename)
    open_pdf_file(leihschein_filename)

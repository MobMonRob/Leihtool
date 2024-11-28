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
from pathlib import Path
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

class FormularData:
    """
    Klasse für alle Formulardaten
    """
    def __init__(self):
        self.studiengang = None
        self.name = None
        self.kurs = None
        self.email = None
        self.ausgeliehene_artikel = [Artikel()]
        self.rueckgabedatum = None
        self.verwendungszweck = None
        self.ausgegeben_durch = None
        self.leihdatum = None
class DefaultValues:
    """
    Klasse für Standardwerte
    """
    def __init__(self):
        self.studiengang = None
        self.verwendungszweck = None
        self.ausgegeben_durch = None

class NameValidator(Validator):
    """
    Validierungsklasse für den Namen
    """
    def validate(self, document: Document) -> None:
        if len(document.text) == 0:
            raise ValidationError(message="Bitte geben Sie einen Namen ein.")
       
class AnzahlValidator(Validator):
    """
    Validierungsklasse für Anzahl der ausgeliehenen Artikel
    """
    def validate(self, document: Document) -> None:
        if not document.text.isdigit() or int(document.text) < 1: 
            raise ValidationError(message="Bitte geben Sie eine Zahl größer 0 ein. Sie füllen gerade einen Leihschein aus, um mindestens einen Artikel zu verleihen.")

class NumberValidator(Validator):
    """
    Validierungsklasse für Zahlen
    """
    def validate(self, document: Document) -> None:
        if not document.text.isdigit(): 
            raise ValidationError(message="Bitte geben Sie eine Zahl ein.")

class EMailValidator(Validator):
    """
    Validierungsklasse für E-Mail-Adressen
    """
    def validate(self, document: Document) -> None:
        if not re.fullmatch(r"[^@]+@[^@]+\.[^@]+", document.text): 
            raise ValidationError(message="Bitte geben Sie eine gültige E-Mail-Adresse ein.")

class ReturnDateValidator(Validator):
    """
    Validierungsklasse für Angabe des Rückgabedatums
    """
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

def get_list_of_studiengaenge() -> list[str]:
    """
    Gibt eine Liste von Studiengängen zurück
    :return: Liste von Studiengängen als Array
    """
    return [
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
        ]

def generate_uniform_leihschein_filename(p_formular_data: FormularData):
    """
    Hilfsmethoden zur Erstellung eines einheitlichen Dateinamens für Leihscheine
    :param p_formular_data: Alle Formulardaten
    :return: Einheitlicher Dateiname im Format <Heutiges Datum>-<Rückgabedatum>-Leihschein-<Name der ausleihenden Person>.pdf
    """
    current_date_yyyy_mm_dd = datetime.now().strftime('%Y-%m-%d')
    rueckgabedatum_datetime = datetime.strptime(p_formular_data.rueckgabedatum, "%d.%m.%Y")
    rueckgabedatum_formatted = rueckgabedatum_datetime.strftime('%Y-%m-%d')

    return f'{current_date_yyyy_mm_dd}-{rueckgabedatum_formatted}-Leihschein-{p_formular_data.name}.pdf'


def generate_leihschein_pdf(p_formular_data: FormularData, p_leihschein_filename):
    """
    Generiert basierend auf einem Template und den eingegebenen Informationen einen ausgefüllten Leihschein
    :param p_formular_data: Alle Formulardaten
    :param p_leihschein_filename: Gewünschter Dateiname des Leihschein-PDFs
    :return: Keine Rückgabe
    """
    # Original PDF template file
    template_file = os.path.join(APPLICATION_PATH, 'Ausleihe_leer.pdf')

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
                    "Studiengang": p_formular_data.studiengang,
                    "Name": p_formular_data.name,
                    "Kurs": p_formular_data.kurs,
                    "Email": p_formular_data.email,
                    "Rückgabedatum": p_formular_data.rueckgabedatum,
                    "Verwendungszweck": p_formular_data.verwendungszweck,
                    "Ausgegeben durch": p_formular_data.ausgegeben_durch,
                    "Datum": p_formular_data.leihdatum
                }
            )
            ausgeliehene_artikel_line = 0
            for p_artikel in p_formular_data.ausgeliehene_artikel:
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


def create_outlook_task_as_reminder(p_formular_data: FormularData):
    """
    Hilfsmethode zur Erstellung eines Outlook Tasks mit Ablaufdatum=Rückgabedatum und allen relevanten Informationen zur Ausleihe
    :param p_formular_data: Alle Formulardaten
    :return:
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    tasks_folder = outlook.getDefaultFolder(13)
    task = tasks_folder.Items.Add()

    ausgeliehene_artikel_body = ''
    for artikel in p_formular_data.ausgeliehene_artikel:
        ausgeliehene_artikel_body += artikel.bezeichnung + ', '

    rueckgabedatum_datetime = datetime.strptime(p_formular_data.rueckgabedatum, "%d.%m.%Y")
    rueckgabedatum_formatted = rueckgabedatum_datetime.strftime('%m/%d/%Y')

    task.Subject = 'Rückgabe des Leihscheins ' + generate_uniform_leihschein_filename(p_formular_data)
    task.Body = f'Die Rückgabe von {p_formular_data.name} ({p_formular_data.email}) aus dem Kurs {p_formular_data.kurs} ist fällig. Folgende Artikel wurden am {p_formular_data.leihdatum} für {p_formular_data.verwendungszweck} geliehen: \n{ausgeliehene_artikel_body}'
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


    absolut_path_to_leihschein_pdf = os.path.join(EXECUTION_PATH, p_leihschein_filename)
    mail.Attachments.Add(absolut_path_to_leihschein_pdf)

    mail.Send()

def show_menu():
    """
    Zeigt ein Menü mit verschiedenen Optionen an
    :return: keine Rückgabe
    """
    menu_options = [
        "Standardwerte festlegen",
        "Einstellungen für externe Anwendungen (Obsidian)"
    ]
    choice = questionary.select(
        "Wählen Sie eine Option:",
        choices=menu_options
    ).ask()

    if choice == "Standardwerte festlegen":
        # Aktion für Option 1
        set_default_values()
    elif choice == "Einstellungen für externe Anwendungen (Obsidian)":
        # Aktion für Option 2
        pass

# Funktion, die auf F1-Taste reagiert
def on_f1_press(event):
    """
    Funktion, die auf F1-Taste reagiert und das Menü anzeigt
    :param event: Event-Objekt
    :return: keine Rückgabe
    """
    sys.stdout.write('\033[2K\033[1G')  # Löscht die aktuelle Zeile im Terminal
    print("F1 wurde gedrückt. Öffne Menü...")
    show_menu()

def set_default_values():
    """
    Setzt Standardwerte für das Leihscheintool basierend auf Benutzereingabe
    """
    default_values = DefaultValues()
    default_values.studiengang = questionary.autocomplete("Studiengang:", choices=get_list_of_studiengaenge()).ask()
    default_values.verwendungszweck = questionary.text("Verwendungszweck:").ask()
    default_values.ausgegeben_durch = questionary.text("Ausgegeben durch:").ask()
    save_default_values(default_values)

def save_default_values(p_default_values: DefaultValues):
    """
    Speichert Standardwerte in einer Konfigurationsdatei
    :param p_default_values: Standardwerte
    :return: keine Rückgabe
    """
    try:
        dir_of_leihschein_settings = os.path.join(USER_PATH, 'LeihscheinTool')
        path_of_default_values = os.path.join(dir_of_leihschein_settings, 'default_values.txt')
        # Erstelle Verzeichnis, falls es noch nicht existiert
        Path(dir_of_leihschein_settings).mkdir(parents=True, exist_ok=True)
        with open(path_of_default_values, 'w', encoding="utf-8") as default_values_file:
            default_values_file.write(p_default_values.studiengang + '\n')
            default_values_file.write(p_default_values.verwendungszweck + '\n')
            default_values_file.write(p_default_values.ausgegeben_durch + '\n')
    except FileNotFoundError:
        print("Standardwerte konnten nicht gespeichert werden.")

def load_default_values() -> DefaultValues:
    """
    Lädt Standardwerte aus einer Konfigurationsdatei
    :return: Standardwerte als Klassenobjekt DefaultValues
    """
    default_values = DefaultValues()
    try:
        with open(os.path.join(USER_PATH, 'LeihscheinTool', 'default_values.txt'), 'r', encoding="utf-8") as default_values_file:
            default_values.studiengang = default_values_file.readline().strip()
            default_values.verwendungszweck = default_values_file.readline().strip()
            default_values.ausgegeben_durch = default_values_file.readline().strip()
    except FileNotFoundError:
        print("Standardwerte konnten nicht geladen werden. Es werden leere Standardwerte verwendet.")
    return default_values

def main():
    """
    Hauptfunktion, die den Ablauf des Leihscheintools definiert
    """
    # Programmheader
    print(f"{'*' * 60}\n Sie verwenden das Leihscheintool in Version {___version__}\n" 
           + " Es hilft Ihnen beim Verleih von Hardware der DHBW.\n"
           + "\n"
           + " Sie können jederzeit mit F1 ins Menü wechseln,\n"
            + " um Einstellungen vorzunehmen.\n"
          + f"{'*' * 60}")
    # Überwachen der F1-Taste
    keyboard.on_press_key("F1", on_f1_press)

    # Standardwerte laden
    default_values = load_default_values()

    formular_data = FormularData()

    # Eingabeaufforderungen für alle Formularfelder
    formular_data.studiengang = questionary.autocomplete(
        'Studiengang:',
        choices=get_list_of_studiengaenge(), default=default_values.studiengang).ask()
    formular_data.name = questionary.text('Name:', validate=NameValidator).ask()
    formular_data.kurs = questionary.text('Kurs:').ask()
    formular_data.email = questionary.text('Email:', validate=EMailValidator).ask()
    anzahl_ausgeliehene_artikel = int(questionary.text('Anzahl ausgeliehene Artikel:', validate=AnzahlValidator).ask())
    formular_data.ausgeliehene_artikel = [Artikel() for i in range(anzahl_ausgeliehene_artikel)]
    for artikel in formular_data.ausgeliehene_artikel:
        questionary.print(f"Eingabe des {formular_data.ausgeliehene_artikel.index(artikel) + 1}. Artikels:", style='bold')
        artikel.pos = int(questionary.text('Pos.:', validate=NumberValidator).ask())
        artikel.menge = int(questionary.text('Menge:', validate=NumberValidator).ask())
        artikel.bezeichnung = questionary.text('Bezeichnung:').ask()
        artikel.seriennummer = questionary.text('Seriennummer:').ask()
        artikel.inventar_nummer = questionary.text('Inventar-Nummer:').ask()
    formular_data.rueckgabedatum = questionary.text('Rückgabedatum:', validate=ReturnDateValidator).ask()
    formular_data.verwendungszweck = questionary.text('Verwendungszweck:', default=default_values.verwendungszweck).ask()
    formular_data.ausgegeben_durch = questionary.text('Ausgegeben durch:', default=default_values.ausgegeben_durch, validate=NameValidator).ask()
    formular_data.leihdatum = datetime.now().strftime('%d.%m.%Y')

    leihschein_filename = generate_uniform_leihschein_filename(formular_data)

    # Generiere Leihschein PDF
    generate_leihschein_pdf(formular_data, leihschein_filename)
    create_outlook_task_as_reminder(formular_data)
    notify_lender_by_email = questionary.confirm("Soll eine E-Mail an die ausleihende Person gesendet werden?").ask()
    if notify_lender_by_email:
        send_email_to_lender(formular_data.email, leihschein_filename)
        print("E-Mail an die ausleihende Person wurde versendet.")
    open_pdf_file(leihschein_filename)


if __name__ == "__main__":
    # Prüfe, wie das Python-Skript ausgeführt wird, um Applikationspfad korrekt zu setzen
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        APPLICATION_PATH = sys._MEIPASS
    else:
        APPLICATION_PATH = os.path.dirname(os.path.abspath(__file__))
    EXECUTION_PATH = os.path.dirname(sys.executable)
    USER_PATH = os.path.expanduser("~")
    main()

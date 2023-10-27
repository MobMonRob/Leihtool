# Leihtool
Python-Skript zur einheitlichen Anlage von Leihscheinen an der DHBW Karlsruhe. Alle Formularfelder werden abgefragt und ausgefüllt.
Zusätzlich wird ein Outlook-Task mit Ablaufdatum angelegt und das generierte PDF zur Überprüfung und anschließenden Druck automatisch geöffnet.

## Abhängigkeiten
- Python=3.9
- pywin32=305.1
- typing_extensions=4.7.0
- pypdf=3.11.1
- pyinstaller=5.6.2

## Bau
Empfohlen ist für den Bau einer Windows-Executable das Tool [PyInstaller](https://pyinstaller.org/en/stable/index.html).
Folgender Befehl baut eine einzige EXE-Datei, die alle benötigten Ressourcen enthält:
```
pyinstaller leihtool.py --onefile --add-data 'Ausleihe_leer.pdf;.'
```

## Template
Das Formular basiert auf dem PDF 'Ausleihe_leer.pdf'.

# Leihtool
Python-Skript zur einheitlichen Anlage von Leihscheinen an der DHBW Karlsruhe. Alle Formularfelder werden abgefragt und ausgefüllt.
Zusätzlich wird ein Outlook-Task mit Ablaufdatum angelegt und das generierte PDF zur Überprüfung und anschließenden Druck automatisch geöffnet.

## Abhängigkeiten
- pywin32
- typing_extensions
- pypdf
- pyinstaller

## Bau
Empfohlen ist für den Bau einer Windows-Executable das Tool [PyInstaller](https://pyinstaller.org/en/stable/index.html).
Folgender Befehl baut eine einzige EXE-Datei, die alle benötigten Resourcen enthält:
```
pyinstaller leihtool --onefile --add-data 'Ausleihe_leer.pdf;.'
```

## Template
Das Formular basiert auf dem PDF 'Ausleihe_leer.pdf'.

# BMA Screenshot Analyzer

Windows-Desktop-Anwendung zur lokalen Auswertung von BMA-Diagnose-Screenshots
und zur Uebertragung der erkannten Ringdaten in eine Excel-Vorlage.

## Wichtiger Lizenzhinweis

**Dies ist keine Open-Source-Software.** Der Quellcode ist nur zur Ansicht
oeffentlich. Nutzung, Ausfuehrung, Kopieren, Veraendern, Weitergabe, Hosting
oder kommerzielle Verwertung sind ohne vorherige ausdrueckliche schriftliche
Erlaubnis des Rechteinhabers untersagt. Massgeblich ist die Datei `LICENSE`.

Drittanbieter-Bibliotheken und PaddleOCR-Modelle behalten ihre eigenen
Lizenzen; Details stehen in `THIRD_PARTY_NOTICES.md`.

## GitHub-Version 1.0.1

- Das grosse rote BMA-Branding im Kopfbereich wurde entfernt.
- OCR und Auswertung laufen weiterhin vollstaendig lokal.
- Der Portable-Build enthaelt die benoetigten Detektions- und
  Erkennungsmodelle.

## Voraussetzungen fuer einen lokalen Build

- Windows x64
- Python 3.13
- PowerShell

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\build-portable.ps1
```

Das Release-ZIP wird unter `release/` erzeugt. Build-Ausgaben werden nicht in
Git eingecheckt.

## Verwendung

Eine Verwendung der Anwendung ist nur mit vorheriger ausdruecklicher
schriftlicher Genehmigung des Rechteinhabers erlaubt. Die technische Anleitung
liegt in `Anleitung.txt` und wird mit dem Release-Paket ausgeliefert.

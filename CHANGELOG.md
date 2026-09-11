# Changelog

Alle nennenswerten Änderungen an diesem Projekt werden hier festgehalten.
Es gibt keine versionierten Releases - das aktuellste `latest`-Release
entspricht immer dem `main`-Branch, siehe [README](README.md#download-fertige-programme-ohne-python-installation).

## [Unreleased]

### Hinzugefügt
- README.md mit Kurzanleitung, Download-Link, Format der
  Mitgliederliste, Kommandozeilen- und Build-Anleitung
- CHANGELOG.md (diese Datei)
- Hinweis zur Windows-SmartScreen-Warnung bei unsignierten Programmen

### Geändert
- GitHub-Actions-Workflow veröffentlicht die gebauten Programme jetzt
  automatisch als GitHub Release (`latest`) statt als schwer auffindbares
  Workflow-Artifact - kein GitHub-Account zum Herunterladen mehr nötig

## 2026-09-11 - Erste Version

### Hinzugefügt
- `sepa_lastschrift.py`: Python-Ersatz für das VBA-Makro
  `Excel2SepaXML.xlsm` - erzeugt SEPA-Lastschriften (pain.008.001.02-XML)
  aus einer Excel-/CSV-Mitgliederliste, inkl. IBAN-Prüfziffer und
  denselben Validierungsregeln wie im Original-Makro
- `gui.py`: grafische Oberfläche (Mitgliederliste auswählen,
  Fälligkeitstag eintragen mit „+7 Tage"-Schnellauswahl, Vereinsdaten
  verwalten) für die Nutzung ohne Kommandozeile
- GitHub-Actions-Workflow zum automatischen Bauen von Windows- und
  macOS-Programmen bei jedem Push

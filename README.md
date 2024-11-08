# Arbeitszeiterfassungssystem

Dieses System hilft bei der Erfassung, Verwaltung und Auswertung von Arbeitsstunden für fünf Mitarbeiter.

## Systemaufbau

1. Excel-Templates:
   - Mitarbeiter-Zeiterfassung
   - Admin-Dashboard

2. VBA-Skript für die Mitarbeiter-Zeiterfassung
3. Python-Skript für die Datenverarbeitung und -analyse

## Einrichtung des Systems

### Schritt 1: Excel-Dateien erstellen

1. Erstellen Sie eine neue Excel-Datei für die Mitarbeiter-Zeiterfassung.
2. Importieren Sie die Daten aus `templates_and_data/mitarbeiter_zeiterfassung_template.csv` in die Excel-Datei.
3. Speichern Sie die Datei als `mitarbeiter_zeiterfassung.xlsx`.

4. Erstellen Sie eine neue Excel-Datei für das Admin-Dashboard.
5. Importieren Sie die Daten aus `templates_and_data/admin_dashboard_template.csv` in die Excel-Datei.
6. Speichern Sie die Datei als `admin_dashboard.xlsx`.

### Schritt 2: VBA-Skript einfügen

1. Öffnen Sie die `mitarbeiter_zeiterfassung.xlsx` Datei.
2. Drücken Sie `Alt + F11`, um den VBA-Editor zu öffnen.
3. Fügen Sie ein neues Modul hinzu (Einfügen > Modul).
4. Kopieren Sie den Inhalt der Datei `mitarbeiter_zeiterfassung_vba.txt` in das neue Modul.
5. Speichern Sie die Excel-Datei als `macro-enabled workbook` (.xlsm).

### Schritt 3: Python-Umgebung einrichten

1. Stellen Sie sicher, dass Python 3.7 oder höher installiert ist.
2. Installieren Sie die erforderlichen Pakete mit:
   ```
   pip install openpyxl
   ```

### Schritt 4: System verwenden

1. Mitarbeiter füllen ihre Arbeitszeiten in der `mitarbeiter_zeiterfassung.xlsx` Datei aus.
2. Der Administrator verwendet das Python-Skript, um die Daten zu verarbeiten und zu analysieren:
   ```
   python adjust_hours.py
   ```
3. Die Ergebnisse werden in der `test_angepasste_zeiterfassung.csv` Datei gespeichert.
4. Der Administrator aktualisiert das `admin_dashboard.xlsx` mit den verarbeiteten Daten.

## Hinweise

- Stellen Sie sicher, dass alle Dateien im gleichen Verzeichnis gespeichert sind.
- Befolgen Sie die österreichischen Arbeitsgesetze bei der Verwendung dieses Systems.
- Kontaktieren Sie den Support bei Fragen oder Problemen.

### Schritt 5: GUI-Anwendung verwenden

1. Starten Sie die GUI-Anwendung mit:
   ```
   python arbeitszeiterfassung_gui.py
   ```
2. Klicken Sie auf "Zeiterfassungsdatei auswählen" und wählen Sie die CSV-Datei mit den Mitarbeiterdaten aus.
3. Klicken Sie auf "Daten verarbeiten", um die Arbeitsstunden anzupassen.
4. Klicken Sie auf "Daten analysieren", um eine Übersicht der angepassten Daten zu erhalten.

Die GUI-Anwendung führt die Datenverarbeitung und -analyse durch, ohne dass Sie direkt mit Python-Skripten oder CSV-Dateien interagieren müssen.


### Schritt 3a: Python-Abhängigkeiten installieren

Installieren Sie alle erforderlichen Python-Pakete mit:

```
pip install -r requirements.txt
```




## Systemanforderungen

- Python 3.7 oder höher
- Alle in der `requirements.txt` aufgeführten Python-Pakete

## Installation

1. Stellen Sie sicher, dass Python 3.7 oder höher installiert ist.
2. Klonen Sie dieses Repository oder laden Sie es herunter.
3. Navigieren Sie im Terminal zum Projektverzeichnis.
4. Installieren Sie die erforderlichen Pakete mit:

   ```
   pip install -r requirements.txt
   ```

## Verwendung der GUI-Anwendung

1. Starten Sie die GUI-Anwendung mit:

   ```
   python arbeitszeiterfassung_gui.py
   ```

2. Klicken Sie auf "Zeiterfassungsdatei auswählen" und wählen Sie die CSV-Datei mit den Mitarbeiterdaten aus.
3. Klicken Sie auf "Daten verarbeiten", um die Arbeitsstunden anzupassen.
4. Klicken Sie auf "Daten analysieren", um eine Übersicht der angepassten Daten zu erhalten.

Die GUI-Anwendung führt die Datenverarbeitung und -analyse durch, ohne dass Sie direkt mit Python-Skripten oder CSV-Dateien interagieren müssen.


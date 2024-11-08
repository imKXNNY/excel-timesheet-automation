# Administrator Guide

## Testschritte:
### **Schritt 1:** Verarbeiten einer Arbeitszeiterfassungsdatei
- **GUI starten:** Führe *main.py* aus, um die GUI zu starten.
```bash
python main.py
```
- **Datei auswählen:**

    - Klicke auf den Button "Zeiterfassungsdatei auswählen".
    - Wähle eine Beispiel-Mitarbeiterdatei (z.B., Mitarbeiter1.xlsx) aus dem Verzeichnis data/employees/.

- **Daten verarbeiten:**

    - Klicke auf "Daten verarbeiten".
    - Das Makro adjust_hours wird ausgeführt, das die Gesamtstunden berechnet und eine neue Datei erstellt (z.B., Mitarbeiter1_angepasst.xlsx).
    - Überprüfe die Log-Datei (arbeitszeiterfassung.log) auf Erfolgsmeldungen.
    
- **Ergebnis überprüfen:**

Öffne die angepasste Datei (Mitarbeiter1_angepasst.xlsx).
Stelle sicher, dass die Arbeitsstunden entsprechend den Anpassungsregeln verändert wurden (z.B., Überstunden korrekt verteilt und auf 15 Minuten gerundet).
Schritt 2: Validieren der Daten
Daten validieren:

Klicke auf den Button "Daten validieren".
Das Makro validate_data wird ausgeführt, das die Daten auf Plausibilität überprüft.
Ergebnis überprüfen:

Überprüfe die Log-Datei (arbeitszeiterfassung.log) auf Validierungsergebnisse.
Falls Fehler oder Warnungen auftreten, werden diese in den Log-Dateien und gegebenenfalls in MessageBoxen angezeigt.
Schritt 3: Analysieren der angepassten Daten
Daten analysieren:

Klicke auf den Button "Daten analysieren".
Die Funktion analyze_adjusted_data wird ausgeführt, die eine Zusammenfassung der Gesamtstunden, Überstunden und Urlaubstage erstellt.
Ergebnis anzeigen:

Ein neues Fenster öffnet sich mit der Analysezusammenfassung.
Verifiziere, dass die Analysewerte den erwarteten Ergebnissen entsprechen.
Schritt 4: Importieren der Mitarbeiterdaten ins Admin-Dashboard
Mitarbeiterdaten importieren:

Klicke auf den Button "Mitarbeiterdaten importieren".
Wähle die Admin-Dashboard-Datei (z.B., admin_dashboard.xlsx) aus.
Wähle das Verzeichnis aus, das die angepassten Mitarbeiterdateien enthält (data/employees/).
Daten importieren:

Die Funktion import_employee_data wird ausgeführt, die die Daten in das Admin-Dashboard importiert und aktualisiert.
Ergebnis überprüfen:

Öffne die Admin-Dashboard-Datei (admin_dashboard.xlsx).
Überprüfe, ob die Daten für jeden Mitarbeiter korrekt aktualisiert wurden, einschließlich Gesamtstunden, Überstunden und Urlaubstage.

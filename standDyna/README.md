# Accesstabellen für standDyna konvertieren und Belegungsdichte erstellen

## Diese Anleitung kann auf DC oder DG Tabellen (im Folgen Datentabelle genannt) angewendet werden.

1. Excel vorbereiten, Entwicklertools aktivieren
⋅⋅1. Klicken Sie auf die Registerkarte Datei.
⋅⋅2. Klicken Sie auf Optionen.
⋅⋅3. Klicken Sie auf Menüband anpassen.
⋅⋅4. Aktivieren Sie unter Menüband anpassen und unter Hauptregisterkarten das Kontrollkästchen Entwicklertools.
2. Access Daten (DC oder DG Tabelle) aus Access exportieren
⋅⋅1. Access Datei öffnen
⋅⋅2. rechtsklick auf die Datentabelle
⋅⋅3. exportieren als "Excel"
3. Excel Daten konvertieren
⋅⋅1. Excel Datei mit Marko (convertAccessHemmenhofenToStandDyna.xlsm) öffnen
⋅⋅2. Exportierte Excel Datei öffnen
⋅.3. Unter "Entwicklertools" auf "Makros" klicken
⋅⋅4. Unter "Markos in" "convertAccessHemmenhofenToStandDyna.xlsm" auswählen
⋅⋅5. "Ausführen" klicken
⋅⋅6. Für standDyna nicht kompatible Kurven mit 0 oder 1 als Wert werden ignoriert und in einer entsprechenden Meldung angezeigt
⋅⋅7. Tabelle mit Daten „Belegungsdichte“ erscheint
⋅⋅8. Belegungsdichtedaten können verwendet werden
⋅⋅9. Tabelle mit Daten für „standDyna“ erscheint
4. Daten für standDyna exportieren
⋅⋅1. Tabelle „standDyna“ auswählen
⋅⋅2. „Datei“ -> „speichern unter“ -> Dateityp: CSV (Trennzeichen-getrennt) (*.csv) -> „speichern“
⋅⋅3. „ok“, dass nur das aktuelle Blatt gespeichert werden soll
⋅⋅4. „ja“, dass Merkmale enthalten sind, die nicht mit .csv kompatibel sind
⋅⋅5. Excel Datei schließen, dabei „nicht speichern“
5. Daten in standDyna einlesen
⋅⋅1. „Stand-Dyna_1-5.exe“ öffnen
⋅⋅⋅⋅1. Ggf. Parameter event und Parameter sustainability anpassen 
⋅⋅⋅⋅2. Unter “Files” -> “Input” die exportierte csv Datei mit “open” öffnen
⋅⋅⋅⋅3. Pfade für „Output events“ und „Output events merged“ ggf. Anpassen
⋅⋅⋅⋅4. Mit “go” Berechnung starten

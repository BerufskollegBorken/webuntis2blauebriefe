# webuntis2BlaueBriefe

Webuntis2BlaueBriefe generiert versandfertige Blaue Briefe als einzelne Worddokumente.

Technische Voraussetzung sind: 
1. der Einsatz von Atlantis als Schulverwaltungsprogramm
2. der Einsatz von Exchange (evtl. auch O365 oder ein anderer Mailservice nach entsprechenden Anpassungen des Codes)
3. der Einsatz von Webuntis und Untis auf Basis einer MDB-Datenbank
4. Installation von Word 

Organisatorische Voraussetzungen sind:
1. eine Prüfungsart namens  "Blauer Brief" ist in Webuntis angelegt
2. die Halbjahresnoten wurden in einer Prüfungsart namens Halbjahreszeugnis angelegt
2. die Lehrerinnen und Lehrer haben die Blauen Briefe in der Prüfungseintrag "Blauer Brief" eingetragen 
3. die Datei ```MarksPerLesson.csv``` wurde aus Webuntis exportiert und auf dem Desktop abgelegt.

## Detailierte Schritte der Erstellung der Datei namens MarksPerLesson.csv:

### Exportieren Sie die Datei aus dem Digitalen Klassenbuch, indem Sie:

1. Klassenbuch > Berichte klicken
2. Alle Klassen auswählen
3. Unter "Noten" die Prüfungsart "Alle" auswählen
4. Hinter "Noten pro Schüler" auf CSV klicken
5. Die Datei "MarksPerLesson.csv" auf dem Desktop speichern.

## Handhabung:

1. Anpassung der Datei Global.
2. Sart des Programms.

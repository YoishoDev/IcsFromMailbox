*Read this in other languages: [English](README.md)*

# IcsFromMailbox
Einen Kalender (Ereignisse) aus einem öffentlichen Ordner Microsoft Exchange System laden und in eine iCalendar Datei umwandeln (PowerShell).

## Idee und Hintergrund
Für eine WordPress-Installtion soll mit Hilfe eines [Plug-Ins](https://de.wordpress.org/plugins/ics-calendar/) ein Kalender aus einen Microsoft Exchange 2019 (On-premises) dargestellt werden.
Hierfür benötigt das Plug-In eine URL auf eine iCalendar-Datei. Da Exchange 2019 keine Möglchkeit bietet, eine Kalender über eine URL zu veröffentlichen, habe ich mir überlegt,
Ereignisse aus einem Kalender des Exchange-Systemes in eine ICS-Datei zu exportieren und diese einfach im Verzeichnis der WordPress-Installation abzulegen.
Da die WordPress-Installation auf einem Microsoft Windows System läuft, sollte es nach Möglichkeit ein PowerShell-Skript sein, damit nicht noch weitere Umgebungen installiert werden müssen.
Ich hatte zwar eine PHP-Variante gefunden, PHP ist ja installiert, um aus einer CSV-Datei eines ICS-Datei zu erzeugen. Ich hätte jedoch immer noch etwas anderes benötigt, umd die Ereignisse aus dem 
Exchange-System laden.
Da ich nichts Passendes (im Ganzen) gefunden habe, hier nun ein Variante mit Kombinantionen verschiedener Dinge, die ich im Netz gefunden habe.

- Export der Ereignisse mittels eines PS-Cmdlets von [hier](https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts) in eine CSV-Datei
- Umwandeln in eine ICS-Datei
- Ablage in der WordPress-Installation

Lange Zeit habe ich damit verbracht, das die erzeugte ICS-Datei die Validätsprüfung bestand. Die Inhalte waren immer ok, nur die Datei "hatte etwas" am Anfang. 
Ich habe bis heute keine richtige Erklärung für das Verhalten :confused:.

**Ein weiteres offenes Problem ist, dass die Ausführung des Skripts nur unter Windows 10 funktioniert, aber nicht auf einem (beliebigen) Serverbetriebsystem. 
Dort wird vom Exchange-Server immer folgender Fehler zurückgegeben: "Der Remoteserver hat einen Fehler zurückgegeben: (401) Nicht autorisiert.".
Im Exchange-Server wird dann "Audit_Failure 4624, 0xc000035b" geloggt.**

Vielleicht dienen die Skripte ja jemanden als Anregung für ähnliche "Probleme" und ersparen damit lange und evtl. erfolglose Suchen.

## Wünsche, Kritik, Fehler, Anmerkungen, Verbesserungsvorschläge ...

Her damit! :+1: Bitte schreiben Sie mir alles.
Entweder Sie nutzen gleiche diese [Seite](https://github.com/YoishoDev/IcsFromMailbox/issues) oder\
Sie schreiben mir eine <a href="mailto:development@yoisho.de">E-Mail<a>.

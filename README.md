# Office Backup System

Ein umfassendes Backup-Tool für Office-Dokumente mit GUI, Speicherplatz-Prüfung und VM-Upload Funktion.

## Allgemeine Informationen
* **Start:** Das Skript kann ausgeführt werden, ohne dass direkt ein Backup gestartet wird (es öffnet sich eine Benutzeroberfläche).
* **Rechte:** Es werden **keine Administrator-Rechte** benötigt, damit das Skript funktioniert.
* **Hintergrund-Prozess:** Wenn das Backup gestartet ist, kann normal weitergearbeitet werden. Man erhält eine Windows-Benachrichtigung, sobald das Backup abgeschlossen ist.

---

## Funktionalitäten

### 1.0 Quellordner
Hier kann ausgewählt werden, welcher Ordner (und dessen Unterordner) gesichert werden soll. Es können mehrere Quellordner hinzugefügt werden.

### 2.0 Zielordner
Hier kann ausgewählt werden, in welchem Ordner das Backup gespeichert werden soll.

### 3.0 Ausnahmen
Hier können Ordnernamen definiert werden, die **nicht** gesichert werden sollen.
* *Wichtig:* Dies ist nützlich, wenn man z.B. einen OneDrive-Ordner eingebunden hat, aber nicht möchte, dass dieser nochmals lokal gesichert wird (Vermeidung von Dopplungen oder Endlosschleifen).

### 4.0 Dateitypen
Hier kann markiert werden, welche Dateitypen genau gesichert werden sollen (z.B. nur Word & Excel, oder alles).

### 5.0 Modus

* **Smart (Auto):** Das System entscheidet selbstständig. Wenn kein Backup existiert oder das letzte Voll-Backup älter als eine Woche ist, wird ein **Full Backup** erstellt. Ansonsten ein inkrementelles.
* **Vollständig:** Erstellt zwingend ein komplettes Backup aller Dateien.
* **Inkrementell:** Speichert nur Dateien, die seit dem letzten vollständigen Backup neu hinzugekommen oder geändert wurden. (Falls kein Basis-Backup existiert, wird automatisch ein vollständiges erstellt).

---

## Weitere Tabs & Funktionen

### Historie
Im Reiter "Historie" kann man unten auf **Aktualisieren** klicken. Daraufhin sieht man eine Liste, wann welches Backup erstellt wurde.

### Tools & VM
Hier kann eine IP-Adresse einer VM (Virtual Machine) oder eines externen Servers eingegeben werden. Dies ermöglicht es, die lokalen Backups nach Fertigstellung automatisch auf die VM hochzuladen (via SSH/SCP).

### Wartung
Der Knopf **"ALTE BACKUPS BEREINIGEN"** dient der Speicherplatz-Optimierung.
* **Funktion:** Er löscht alle alten Backups, **ausser** das aktuellste Full-Backup und alle Inkremental-Backups, die danach gemacht wurden (die aktuelle Sicherungskette bleibt erhalten).

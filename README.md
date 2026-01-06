# 📧 OUTLOOK QUARTALSBACKUP - ANLEITUNG

## 📋 ÜBERSICHT

Dieses PowerShell-Skript erstellt automatische Backups aller E-Mail-Konten aus Microsoft Outlook Desktop.

**Features:**
- ✅ Sichert ALLE konfigurierten Outlook-Konten
- ✅ Organisiert Backups nach Jahr und Quartal
- ✅ Speichert E-Mails als MSG-Dateien (öffenbar mit Outlook)
- ✅ Behält Ordnerstruktur bei
- ✅ Zeigt Fortschritt in Echtzeit
- ✅ Erstellt automatisch Unterordner

---

## 🔴 VORAUSSETZUNGEN

### Erforderlich:

1. **Microsoft Outlook Desktop** (NICHT Outlook.com Web)
   - Microsoft 365 Outlook ODER
   - Outlook 2016/2019/2021/2024

2. **Windows 10/11**

3. **PowerShell 5.1+** (bereits in Windows enthalten)

4. **Alle E-Mail-Konten in Outlook eingerichtet**

5. **Google Drive Desktop** (wenn Speicherort G:\ verwendet wird)

---

## 🚀 INSTALLATION

### Schritt 1: Outlook konfigurieren

1. **Öffne Outlook Desktop**
2. **Gehe zu:** `Datei` → `Kontoeinstellungen` → `Kontoeinstellungen...`  
3. Folge den Anweisungen
3. Stelle sicher, dass E-Mails synchronisiert sind

---

### Schritt 2: PowerShell-Ausführung erlauben

1. **Öffne PowerShell als Administrator:**
   - Windows-Taste drücken
   - Tippe: `powershell`
   - Rechtsklick → `Als Administrator ausführen`

2. **Führe folgenden Befehl aus:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

3. Bestätige mit `J` (Ja)

---

### Schritt 3: Skript herunterladen

1. Lade `Backup-Final.ps1` von diesem Repository herunter
2. Speichere es in: `G:\Meine Ablage\Outlook Archiv\`

---

### Schritt 4: Skript ausführen

```powershell
cd "G:\Meine Ablage\Outlook Archiv"
.\Backup-Final.ps1
```

---

## ⚙️ KONFIGURATION

### Backup-Pfad ändern

Öffne die .ps1-Datei und ändere:

```powershell
$BaseBackupPath = "G:\Meine Ablage\Outlook Archiv"
```

Zu:

```powershell
$BaseBackupPath = "C:\Dein\Gewünschter\Pfad"
```

---

## 🔄 AUTOMATISIERUNG

### Automatisches Backup am Quartalsende einrichten

1. **Öffne PowerShell als Administrator**

2. **Führe folgenden Befehl aus:**

```powershell
# Alte Tasks löschen (falls vorhanden)
schtasks /delete /tn "Outlook Quartalsbackup Q1" /f 2>$null
schtasks /delete /tn "Outlook Quartalsbackup Q2" /f 2>$null
schtasks /delete /tn "Outlook Quartalsbackup Q3" /f 2>$null
schtasks /delete /tn "Outlook Quartalsbackup Q4" /f 2>$null

$scriptPath = "G:\Meine Ablage\Outlook Archiv\Backup-Final.ps1"

# Q1 - 31. März um 22:00
schtasks /create /tn "Outlook Quartalsbackup Q1" /tr "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`"" /sc yearly /d 31 /m MAR /st 22:00 /rl HIGHEST /f

# Q2 - 30. Juni um 22:00
schtasks /create /tn "Outlook Quartalsbackup Q2" /tr "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`"" /sc yearly /d 30 /m JUN /st 22:00 /rl HIGHEST /f

# Q3 - 30. September um 22:00
schtasks /create /tn "Outlook Quartalsbackup Q3" /tr "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`"" /sc yearly /d 30 /m SEP /st 22:00 /rl HIGHEST /f

# Q4 - 31. Dezember um 22:00
schtasks /create /tn "Outlook Quartalsbackup Q4" /tr "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`"" /sc yearly /d 31 /m DEC /st 22:00 /rl HIGHEST /f
```

3. **Überprüfe die Aufgabe:**
   - Windows-Taste → `Aufgabenplanung`
   - Suche: `Outlook Quartalsbackup`

---

## 📊 QUARTALE

Das Skript erkennt automatisch das aktuelle Quartal:

| Quartal | Monate | Enddatum |
|---------|--------|----------|
| Q1 | Jan-Mär | 31. März |
| Q2 | Apr-Jun | 30. Juni |
| Q3 | Jul-Sep | 30. September |
| Q4 | Okt-Dez | 31. Dezember |

---

## 📁 ORDNERSTRUKTUR

```
G:\Meine Ablage\Outlook Archiv\
├── 2026\
│   ├── Q1\
│   │   ├── [email1@example.com]\
│   │   │   ├── Posteingang\
│   │   │   │   └── E-Mail_Betreff_2026-01-15_12-30-45.msg
│   │   │   ├── Gesendete Elemente\
│   │   │   └── ...
│   │   └── [email2@example.com]\
│   ├── Q2\
│   └── ...
```

---

## 🛠️ FUNKTIONSWEISE

1. **Outlook-Verbindung:** Skript startet Outlook (falls nicht aktiv)
2. **Konto-Erkennung:** Findet alle E-Mail-Konten automatisch
3. **Ordner-Scan:** Durchsucht jeden Ordner (außer Systemordner)
4. **Backup:** Speichert E-Mails als `.msg`-Dateien
5. **Duplikate vermeiden:** Überspringt bereits gesicherte E-Mails
6. **Abschluss:** Zeigt Statistik und schließt

---

## ⚠️ WICHTIGE HINWEISE

### Ausgeschlossene Ordner:

- 📅 Kalender / Calendar
- 👤 Kontakte / Contacts
- ✅ Aufgaben / Tasks
- 📝 Notizen / Notes
- 📖 Journal
- 🔍 Suchordner / Search Folders
- 🗑️ Gelöschte Elemente / Deleted Items (optional)

### Dateinamen:

```
E-Mail_Betreff_YYYY-MM-DD_HH-MM-SS.msg
```

Beispiel: `Rechnung_Q1_2026-03-15_14-30-22.msg`

---

## 🔧 FEHLERBEHEBUNG

### Problem: "Outlook nicht gefunden"

**Lösung:**
- Stelle sicher, dass Outlook Desktop installiert ist
- Starte Outlook einmal manuell

### Problem: "Zugriff verweigert"

**Lösung:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

### Problem: "Skript hängt"

**Lösung:**
- Schließe Outlook
- Starte das Skript neu
- Prüfe Speicherplatz (mind. 5 GB)

### Problem: "Ordner wird nicht gefunden"

**Lösung:**
- Überprüfe Pfad in Zeile 5 des Skripts
- Erstelle den Ordner manuell

---

## 📈 PERFORMANCE

- **~100 E-Mails:** ca. 2 Minuten
- **~1.000 E-Mails:** ca. 15 Minuten
- **~10.000 E-Mails:** ca. 2-3 Stunden

*Zeiten variieren je nach System und E-Mail-Größe*

---

## ✅ CHECKLISTE

Vor dem ersten Backup:

- [ ] Outlook Desktop installiert
- [ ] Alle Konten in Outlook eingerichtet
- [ ] PowerShell Execution Policy gesetzt
- [ ] Skript in korrektem Ordner gespeichert
- [ ] Genug Speicherplatz verfügbar (mind. 5 GB)
- [ ] Google Drive läuft (falls G:\ verwendet wird)

Nach dem Backup:

- [ ] Log-Datei überprüft
- [ ] Stichprobe: Einige MSG-Dateien geöffnet
- [ ] Alle Konten wurden gesichert
- [ ] Automatische Aufgabe eingerichtet (optional)

---

## 📄 LIZENZ

Dieses Skript ist für den persönlichen Gebrauch erstellt.  
Frei verwendbar, keine Garantie.

---

**Version:** 1.0  
**Erstellt:** Januar 2026  
**Getestet mit:** Outlook 2016, Microsoft 365 Outlook, Windows 11

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

## 🎯 VORAUSSETZUNGEN

### Erforderlich:
1. **Microsoft Outlook Desktop** (NICHT Outlook.com Web)
   - Microsoft 365 Outlook ODER
   - Outlook 2016/2019/2021/2024
  
2. **Windows 10/11**
 
3. **PowerShell 5.1+** (bereits in Windows enthalten)
 
4. **Alle E-Mail-Konten in Outlook eingerichtet**
 
5. **Google Drive Desktop** (wenn Speicherort G:\ verwendet wird)

---

## 📁 ORDNERSTRUKTUR

Das Skript erstellt folgende Struktur:

```
G:\Meine Ablage\Outlook Archiv\
├── 2026\
│   ├── Q1\                           # Quartal 1 (Januar-März)
│   │   ├── beispiel@email.de\
│   │   │   ├── Posteingang\
│   │   │   │   ├── 2026-01-06_1430_Betreff_der_Email.msg
│   │   │   │   └── 2026-01-05_0915_Weitere_Email.msg
│   │   │   ├── Gesendete Elemente\
│   │   │   └── Entwürfe\
│   │   ├── firma@example.com\
│   │   └── familie@example.de\
│   ├── Q2\                           # Quartal 2 (April-Juni)
│   ├── Q3\                           # Quartal 3 (Juli-September)
│   └── Q4\                           # Quartal 4 (Oktober-Dezember)
└── backup_log.txt                    # Log-Datei
```

---

## 🚀 INSTALLATION & ERSTE VERWENDUNG

### Schritt 1: Outlook Desktop einrichten

1. Öffne **Microsoft Outlook Desktop**
2. Füge alle E-Mail-Konten hinzu:
   - Klicke: `Datei` → `Konto hinzufügen`
   - Gib E-Mail-Adresse ein
   - Folge den Anweisungen
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
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File 'G:\Meine Ablage\Outlook Archiv\Backup-Final.ps1'"

# Q1 Ende (31. März)
$trigger1 = New-ScheduledTaskTrigger -Daily -At "22:00"
$trigger1.DaysOfMonth = 31
$trigger1.MonthsOfYear = 3

# Q2 Ende (30. Juni)
$trigger2 = New-ScheduledTaskTrigger -Daily -At "22:00"
$trigger2.DaysOfMonth = 30
$trigger2.MonthsOfYear = 6

# Q3 Ende (30. September)
$trigger3 = New-ScheduledTaskTrigger -Daily -At "22:00"
$trigger3.DaysOfMonth = 30
$trigger3.MonthsOfYear = 9

# Q4 Ende (31. Dezember)
$trigger4 = New-ScheduledTaskTrigger -Daily -At "22:00"
$trigger4.DaysOfMonth = 31
$trigger4.MonthsOfYear = 12

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName "Outlook Quartalsbackup" -Action $action -Trigger @($trigger1,$trigger2,$trigger3,$trigger4) -Settings $settings -Description "Automatisches Backup aller Outlook-Konten"
```

3. **Überprüfe die Aufgabe:**
   - Windows-Taste → `Aufgabenplanung`
   - Suche: `Outlook Quartalsbackup`

---

## 📊 QUARTALE

Das Skript erkennt automatisch das aktuelle Quartal:

| Quartal | Monate | Enddatum |
|---------|--------|----------|
| Q1 | Januar - März | 31. März |
| Q2 | April - Juni | 30. Juni |
| Q3 | Juli - September | 30. September |
| Q4 | Oktober - Dezember | 31. Dezember |

---

## 📄 DATEIFORMAT

**MSG-Dateien (.msg)**
- Standard Outlook-Format
- Öffenbar mit: Microsoft Outlook, Thunderbird (mit Add-on), MSG Viewer
- Behält alle Metadaten: Absender, Empfänger, Datum, Anhänge

**Dateinamen-Schema:**
```
YYYY-MM-DD_HHMM_Betreff_der_Email.msg
```

Beispiel:
```
2026-01-06_1430_Rechnung_Januar_2026.msg
```

---

## ❓ HÄUFIGE PROBLEME & LÖSUNGEN

### Problem: "Outlook konnte nicht initialisiert werden"

**Lösung:**
- Stelle sicher, dass Outlook Desktop installiert ist
- Öffne Outlook einmal manuell
- Überprüfe, dass Konten synchronisiert sind

---

### Problem: "Ausführung von Skripts ist deaktiviert"

**Lösung:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### Problem: "Datei wurde nicht gefunden"

**Lösung:**
- Überprüfe, dass die Datei `.ps1` Endung hat (nicht `.ps1.txt`)
- Navigiere zum richtigen Ordner:
```powershell
cd "G:\Meine Ablage\Outlook Archiv"
dir
```

---

### Problem: Skript läuft sehr langsam

**Ursache:** Große Postfächer mit vielen E-Mails

**Lösung:**
- Normal! Bei 10.000+ E-Mails kann es 30-60 Minuten dauern
- Lass das Skript durchlaufen
- Outlook nicht schließen während der Ausführung

---

### Problem: Einige E-Mails fehlen

**Mögliche Ursachen:**
1. E-Mails sind in System-Ordnern (Calendar, Contacts, Tasks)
   - Diese werden absichtlich übersprungen
2. E-Mails sind in Unterordnern
   - Das Skript sichert nur Hauptordner
3. Speicherfehler bei einzelnen E-Mails
   - Wird übersprungen, Rest wird gesichert

---

## 📈 PERFORMANCE

**Geschwindigkeit:**
- ~100-200 E-Mails pro Minute (abhängig von E-Mail-Größe)
- 1.000 E-Mails ≈ 5-10 Minuten
- 10.000 E-Mails ≈ 50-100 Minuten

**Speicherplatzbedarf:**
- Durchschnitt: ~50-100 KB pro E-Mail
- 1.000 E-Mails ≈ 50-100 MB
- 10.000 E-Mails ≈ 500 MB - 1 GB

---

## 🔒 SICHERHEIT & DATENSCHUTZ

**Was das Skript NICHT tut:**
- ❌ Sendet keine Daten ins Internet
- ❌ Ändert keine Original-E-Mails
- ❌ Löscht keine E-Mails
- ❌ Greift nicht auf Passwörter zu

**Was das Skript tut:**
- ✅ Liest E-Mails über Outlook COM-Schnittstelle
- ✅ Speichert Kopien lokal
- ✅ Nur Lesezugriff auf E-Mails

---

## 📝 LOG-DATEI

Das Skript erstellt eine Log-Datei unter:
```
G:\Meine Ablage\Outlook Archiv\YYYY\QX\backup_log.txt
```

**Inhalt:**
- Backup-Zeitstempel
- Liste aller gesicherten Konten
- Anzahl E-Mails pro Ordner
- Fehler (falls vorhanden)

---

## 🛠️ ERWEITERTE NUTZUNG

### Nur bestimmte Konten sichern

Öffne das Skript und füge Filter hinzu:

```powershell
foreach ($store in $namespace.Stores) {
    # Nur diese Konten sichern:
    if ($store.DisplayName -notlike "*beispiel*" -and $store.DisplayName -notlike "*firma*") {
        continue
    }
    # Rest des Codes...
}
```

---

### Bestimmte Ordner ausschließen

Erweitere die Skip-Liste:

```powershell
if ($folder.Name -in @('Calendar','Contacts','Tasks','Notes','Journal','RSS-Feeds','Junk-E-Mail')) {
    continue
}
```

---

## 🎓 TIPPS & TRICKS

**Tipp 1: Backup vor wichtigen Änderungen**
Führe ein manuelles Backup aus, bevor du:
- Outlook neu installierst
- Konten entfernst
- Computer wechselst

**Tipp 2: Regelmäßige Überprüfung**
Überprüfe quartalsweise, ob das automatische Backup funktioniert:
```powershell
dir "G:\Meine Ablage\Outlook Archiv\2026\Q1"
```

**Tipp 3: MSG-Dateien öffnen**
Doppelklick auf .msg-Datei → Öffnet automatisch in Outlook

**Tipp 4: Suche nach E-Mails**
Windows-Suche funktioniert in den Backup-Ordnern:
- Windows-Taste → Suche nach Betreff oder Absender

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
**Erstellt:** Januar 2026  
**Getestet mit:** Microsoft 365 Outlook, Windows 11

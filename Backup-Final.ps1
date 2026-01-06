# ============================================================
# OUTLOOK BACKUP - Alle E-Mail-Konten sichern
# ============================================================
# Erstellt: Januar 2026
# Funktion: Sichert ALLE Outlook-Konten nach Jahr/Quartal
# Format: MSG-Dateien (mit Outlook öffenbar)
# ============================================================

# Basis-Backup-Pfad (Google Drive)
$BaseBackupPath = "G:\Meine Ablage\Outlook Archiv"

# Aktuelles Jahr und Quartal
$Year = (Get-Date).Year
$Month = (Get-Date).Month
$Quarter = [math]::Ceiling($Month / 3)

# Backup-Pfad: Jahr\QX
$BackupPath = Join-Path $BaseBackupPath "$Year\Q$Quarter"

Write-Host "" -ForegroundColor Cyan
Write-Host "=== OUTLOOK BACKUP ===" -ForegroundColor Cyan
Write-Host "Jahr: $Year | Quartal: Q$Quarter" -ForegroundColor White
Write-Host "Ziel: $BackupPath" -ForegroundColor White
Write-Host "" -ForegroundColor Cyan

# Erstelle Backup-Ordner
if (-not (Test-Path $BackupPath)) {
    New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
    Write-Host "[+] Backup-Ordner erstellt" -ForegroundColor Green
}

try {
    # Outlook-Verbindung
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    Write-Host "[+] Outlook verbunden" -ForegroundColor Green
    Write-Host "" -ForegroundColor White
    
    # Durchlaufe ALLE E-Mail-Konten
    foreach ($store in $namespace.Stores) {
        # Kontoname bereinigen
        $accountName = $store.DisplayName -replace '[<>:"/\\|?*]', '_'
        
        # Konto-Backup-Pfad
        $accountPath = Join-Path $BackupPath $accountName
        
        if (-not (Test-Path $accountPath)) {
            New-Item -ItemType Directory -Path $accountPath -Force | Out-Null
        }
        
        Write-Host "Konto: $accountName" -ForegroundColor Yellow
        
        # Hole Root-Folder des Kontos
        $rootFolder = $store.GetRootFolder()
        
        # Durchlaufe alle Ordner im Konto
        foreach ($folder in $rootFolder.Folders) {
            # Ordner-Name
            $folderName = $folder.Name
            
            # Ordnernamen bereinigen
            $folderName = $folderName -replace '[<>:"/\\|?*]', '_'
            
            # Überspringe System-Ordner
            if ($folder.Name -in @('Calendar','Contacts','Tasks','Notes','Journal')) {
                continue
            }
            
            Write-Host "  └─ $folderName" -ForegroundColor Cyan -NoNewline
            
            # Erstelle Ordner-Pfad
            $folderPath = Join-Path $accountPath $folderName
            if (-not (Test-Path $folderPath)) {
                New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
            }
            
            $emailCount = 0
            
            # Sichere alle E-Mails im Ordner
            foreach ($mail in $folder.Items) {
                try {
                    # Nur E-Mails (MailItem)
                    if ($mail.Class -eq 43) {
                        # Betreff bereinigen
                        $subject = if ($mail.Subject) { 
                            $mail.Subject -replace '[<>:"/\\|?*]', '_' 
                        } else { 
                            "Kein_Betreff" 
                        }
                        
                        # Dateiname: YYYY-MM-DD_HHMM_Betreff.msg
                        $receivedTime = $mail.ReceivedTime
                        $timestamp = $receivedTime.ToString("yyyy-MM-dd_HHmm")
                        $fileName = "${timestamp}_${subject}.msg"
                        
                        # Kürze zu lange Dateinamen
                        if ($fileName.Length -gt 200) {
                            $fileName = $fileName.Substring(0, 197) + ".msg"
                        }
                        
                        $filePath = Join-Path $folderPath $fileName
                        
                        # Speichere als MSG
                        $mail.SaveAs($filePath, 3)  # 3 = olMSG
                        $emailCount++
                    }
                } catch {
                    # Fehler überspringen
                }
            }
            
            Write-Host " ($emailCount E-Mails)" -ForegroundColor White
        }
        
        Write-Host "" -ForegroundColor White
    }
    
    Write-Host "[+] Backup abgeschlossen!" -ForegroundColor Green
    Write-Host "" -ForegroundColor White
    
} catch {
    Write-Host "[!] Fehler: $_" -ForegroundColor Red
} finally {
    # Outlook-Objekt freigeben
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Druecke eine Taste zum Beenden..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

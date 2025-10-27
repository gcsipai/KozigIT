# ----------------------------------------------------------------------------------
# Chocolatey Alap- és Sablon Szoftver Telepítő Szkript v1.13 (2025.10)
# Cél: Telepíteni VAGY frissíteni a kért alapvető szoftvereket (upgrade parancs)
# A telepítési lista: 7zip, Google Chrome, Sumatra PDF, Oracle JRE 8, VC Redist, .NET
# Futtatás: Rendszergazdai PowerShell ablakban.
# ----------------------------------------------------------------------------------

# ----------------------------------------------------------------------------------
# 1. Előfeltétel ellenőrzések és Chocolatey telepítése 
# ----------------------------------------------------------------------------------

# Rendszergazdai jogosultság ellenőrzése
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "A szkript futtatásához rendszergazdai jogosultság szükséges!" -ForegroundColor Red
    exit 1
}

Write-Host "Chocolatey telepítő és FRISSÍTŐ szkript v1.13 (Oracle JRE 8 Runtime)" -ForegroundColor Cyan
Write-Host "=========================================================================================" -ForegroundColor Cyan

# Chocolatey telepítése, ha még nem található
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    Write-Host "A Chocolatey nem található. Telepítés megkezdése..." -ForegroundColor Yellow
    
    Set-ExecutionPolicy Bypass -Scope Process -Force
    
    try {
        $script = "https://community.chocolatey.org/install.ps1"
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString($script))
        Write-Host "A Chocolatey sikeresen telepítve!" -ForegroundColor Green
    }
    catch {
        Write-Host "Kritikus hiba történt a Chocolatey telepítése során: $_" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Frissítjük a PATH változót..." -ForegroundColor Yellow
    $env:PATH = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
} else {
    Write-Host "A Chocolatey már telepítve van." -ForegroundColor Cyan
    Write-Host "Chocolatey frissítése..." -ForegroundColor Yellow
    # A Chocolatey saját magát is upgrade-eli
    choco upgrade chocolatey -y
}

# ----------------------------------------------------------------------------------
# 2. Chocolatey konfiguráció 
# ----------------------------------------------------------------------------------

Write-Host "`nChocolatey konfiguráció beállítása automatizálthoz..." -ForegroundColor Yellow
choco feature enable -n allowGlobalConfirmation
choco feature enable -n allowEmptyChecksums

# ----------------------------------------------------------------------------------
# 3. SZOFTVEREK TÖMEGES FRISSÍTÉSE VAGY TELEPÍTÉSE (Csak JRE8)
# ----------------------------------------------------------------------------------

Write-Host "`nSzoftverek frissítése vagy telepítése Chocolatey segítségével..." -ForegroundColor Yellow

# Alapcsomagok listája (JRE8 közvetlenül hozzáadva)
$packages = @(
    "7zip",             # 7-Zip archíváló
    "googlechrome",     # Google Chrome webböngésző
    "sumatrapdf",       # Sumatra PDF olvasó (könnyű és gyors)
    "jre8",             # Oracle JRE 8 Runtime 
    "vcredist-all",     # Visual C++ Redistributable (2015-2022)
    "dotnetfx",         # .NET Framework (telepíti a legújabb 4.8.1-et)
    "dotnet-runtime"    # Legújabb .NET Runtime (pl. .NET 8.0)
)

Write-Host "`nAz Oracle JRE 8 Runtime ('jre8') csomag telepítése automatikusan megtörténik." -ForegroundColor Yellow

# ----------------------------------------------------------------------------------
# VÉGLEGES CSOMAGLISTA MEGJELENÍTÉSE TELEPÍTÉS ELŐTT
# ----------------------------------------------------------------------------------
Write-Host "`n--- VÉGLEGES TELEPÍTÉSI/FRISSÍTÉSI LISTA ---" -ForegroundColor DarkCyan
$packages | ForEach-Object {
    Write-Host "  ✅ $_" -ForegroundColor DarkCyan
}
Write-Host "----------------------------------------------`n" -ForegroundColor DarkCyan

# Csomagok telepítése/frissítése
$needsReboot = $false
foreach ($package in $packages) {
    Write-Host "`nFrissítés/Telepítés: $package" -ForegroundColor Magenta
    try {
        # A choco upgrade telepíti, ha nem létezik, és frissíti, ha létezik.
        choco upgrade $package -y --force --retry 3

        # Visszatérési kód ellenőrzése
        if ($LASTEXITCODE -eq 0) {
            Write-Host "$package sikeresen telepítve/frissítve" -ForegroundColor Green
        } elseif ($LASTEXITCODE -eq 3010) {
            Write-Host "$package telepítve/frissítve, de rendszerújraindítás szükséges" -ForegroundColor Yellow
            $needsReboot = $true
        } elseif ($LASTEXITCODE -eq 1641) {
             # 1641 = sikeres telepítés, azonnali újraindítás függőben
            Write-Host "$package telepítve, újraindítás szükséges" -ForegroundColor Yellow
            $needsReboot = $true
        } else {
            Write-Host "Figyelmeztetés: $package telepítése nem szabványos (exit code: $LASTEXITCODE)." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Hiba: $package telepítése/frissítése során hiba történt: $_" -ForegroundColor Red
        Write-Host "A telepítés folytatódik a következő csomaggal..." -ForegroundColor Yellow
    }
}

# ----------------------------------------------------------------------------------
# 4. TELEPÍTÉS UTÁNI ELLENŐRZÉSEK ÉS JAVÁLLAPOT
# ----------------------------------------------------------------------------------

Write-Host "`nTelepítés utáni ellenőrzések..." -ForegroundColor Yellow

# Telepített csomagok listázása
Write-Host "`nTelepített csomagok:" -ForegroundColor Cyan
choco list --local

# Java ellenőrzés
Write-Host "`nJava állapot ellenőrzése:" -ForegroundColor Cyan
try {
    # Próbáljuk ki a Java parancsot és a hiba kimenetet is elkapjuk
    $javaVersion = cmd /c "java -version 2>&1"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Java telepítve és elérhető:" -ForegroundColor Green
        $javaVersion | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
    } else {
        Write-Host "Java nem található vagy nincs a PATH-ban (hiba kód: $LASTEXITCODE)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Java nem érhető el" -ForegroundColor Yellow
}

# Rendszerújraindítás szükségességének ellenőrzése
if (Test-Path "$env:ProgramData\chocolatey\choco.exe.pending" -or $needsReboot) {
    Write-Host "`n❌ FIGYELMEZTETÉS: Rendszerújraindítás szükséges a telepítések befejezéséhez!" -ForegroundColor Red
    $pendingReboot = $true
} else {
    Write-Host "`n✅ Nincs szükség azonnali rendszerújraindításra" -ForegroundColor Green
    $pendingReboot = $false
}

# ----------------------------------------------------------------------------------
# 5. BEFEJEZÉS
# ----------------------------------------------------------------------------------

Write-Host "`n*** Telepítés befejezve! ***" -ForegroundColor Green

if ($pendingReboot) {
    Write-Host "`n🚨 JAVASLAT: Indítsd újra a számítógépet a telepítések befejezéséhez!" -ForegroundColor Red
}

Write-Host "`nTovábbi hasznos parancsok:" -ForegroundColor Cyan
Write-Host "  - choco upgrade all -y (összes csomag frissítése)" -ForegroundColor Gray
Write-Host "  - choco outdated (elavult csomagok listázása)" -ForegroundColor Gray
Write-Host "  - choco list --local (telepített csomagok listája) " -ForegroundColor Gray

# Megbízható várakozás
Write-Host "`n*** Nyomj Entert az ablak bezárásához. ***" -ForegroundColor Yellow
Read-Host

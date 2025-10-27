# 🚀 Windows 11 Alap Telepítő Szkript (v1.13) 💾

## `windows11-ce.1.13b-alap.ps1`

[![License: Unlicensed](https://img.shields.io/badge/License-Unlicensed-lightgrey.svg)](https://unlicense.org/)
[![PowerShell](https://img.shields.io/badge/Shell-PowerShell-5391FE?logo=powershell&logoColor=white)](https://docs.microsoft.com/en-us/powershell/)
[![Windows 11](https://img.shields.io/badge/Platform-Windows%2011-0078D4?logo=windows&logoColor=white)](https://www.microsoft.com/hu-hu/windows/windows-11)
[![Package Manager](https://img.shields.io/badge/Package%20Manager-Chocolatey-795548?logo=chocolatey&logoColor=white)](https://chocolatey.org/)
[![Java](https://img.shields.io/badge/Java-JRE%208-007396?logo=openjdk&logoColor=white)]()
[![.NET](https://img.shields.io/badge/.NET-Runtime-512BD4?logo=dotnet&logoColor=white)]()

---

## 💡 Áttekintés

Ez a PowerShell szkript egy automata telepítő/frissítő segédprogram, amely a **Chocolatey** csomagkezelő segítségével gondoskodik a Windows 11 alapvető szoftvereinek és futtatókörnyezeti komponenseinek telepítéséről vagy frissítéséről. A szkript a robusztusabb és naprakész telepítés érdekében minden szoftverre a **`choco upgrade`** parancsot használja. A konfiguráció fixen tartalmazza az **Oracle JRE 8 Runtime (`jre8`)** csomagot.

**Verzió:** `v1.13 (Oracle JRE 8 Runtime)`

---

## 💻 Támogatott Platformok és Főbb Szolgáltatások

| Kategória | Alkalmazás / Rendszer | Verzió / Eszköz | Szerep |
| :--- | :--- | :--- | :--- |
| **Operációs Rendszer** | Windows | [![Windows 11](https://img.shields.io/badge/Windows-11-0078D4?logo=windows&logoColor=white)](https://www.microsoft.com/hu-hu/windows/windows-11) | Fő Célplatform |
| **Csomagkezelő** | Chocolatey | [![Chocolatey](https://img.shields.io/badge/Choco-Manager-795548?logo=chocolatey&logoColor=white)](https://chocolatey.org/) | Telepítési és frissítési motor |
| **Futtatókörnyezet** | Oracle JRE 8 | [![Java](https://img.shields.io/badge/Java-8-007396?logo=java&logoColor=white)]() | Java 8 Runtime környezet |
| **Futtatókörnyezet** | .NET Runtime | [![.NET](https://img.shields.io/badge/.NET-Runtime-512BD4?logo=dotnet&logoColor=white)]() | Modern és Framework kompatibilitás |

---

## ✨ Telepített Csomagok és Főbb Funkciók

A szkript a következő alapvető szoftvereket és rendszerkomponenseket telepíti/frissíti:

| Szoftver neve | Chocolatey csomagnév | Leírás |
| :--- | :--- | :--- |
| **Tömörítő** | `7zip` | 7-Zip archíváló és kicsomagoló segédprogram. |
| **Webböngésző** | `googlechrome` | Google Chrome webböngésző. |
| **PDF-olvasó** | `sumatrapdf` | Könnyű, gyors, minimalista PDF-olvasó. |
| **Java Runtime** | `jre8` | Oracle JRE 8 Runtime Environment. |
| **VC Redist** | `vcredist-all` | Microsoft Visual C++ futtatókörnyezet (2015-2022). |
| **.NET Framework** | `dotnetfx` | .NET Framework (legújabb, pl. 4.8.1). |
| **.NET Runtime** | `dotnet-runtime` | Legújabb .NET futtatókörnyezet (pl. .NET 8.0). |

---

## 🚀 Használat

### 1. Előkészítés

Győződjön meg róla, hogy a szkriptet **rendszergazdai jogosultsággal** futtatja, különben a telepítés sikertelen lesz.

### 2. A Szkript Futtatása

1.  **Nyiss Rendszergazdai PowerShellt:** Nyomja le a <kbd>Windows</kbd> gombot, gépelje be: `PowerShell`, majd kattintson jobb egérgombbal a találatra és válassza a **„Futtatás rendszergazdaként”** opciót.
2.  **Másolás és Beillesztés:** Másolja ki a teljes szkriptkódot az alábbi **`PowerShell Szkript Kódja`** szekcióból. Illessze be a kódblokkot a PowerShell ablakba (jobb egérgombbal vagy <kbd>Ctrl</kbd>+<kbd>V</kbd>).
3.  **Futtatás:** Nyomjon <kbd>Enter</kbd>-t.
4.  **Figyeld a futást:** A szkript automatikusan elvégzi a Chocolatey telepítését/frissítését, majd sorban telepíti/frissíti a felsorolt csomagokat.
5.  **Befejezés:** A folyamat végén a szkript ellenőrzi a rendszer újraindítási igényét és a Java állapotát. Nyomjon <kbd>Enter</kbd>-t az ablak bezárásához.

---

## 📜 PowerShell Szkript Kódja (`windows11-ce.1.13b-alap.ps1`)

```powershell
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
        $script = "[https://community.chocolatey.org/install.ps1](https://community.chocolatey.org/install.ps1)"
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

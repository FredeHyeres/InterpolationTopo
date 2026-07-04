# ==============================================================================
# Interpolation Topo - installation automatique (MicroStation V8i SS3)
#
# Copie le projet VBA (.mvba) et la boite a outils (.dgnlib) au bon endroit,
# puis configure le chargement automatique du projet au demarrage.
#
# Usage : clic droit sur install.cmd > Executer  (ou .\install.ps1 en PowerShell)
# Relançable sans risque : chaque etape est ignoree si deja faite.
# ==============================================================================

$ErrorActionPreference = "Stop"
Write-Host "=== Installation Interpolation Topo ===" -ForegroundColor Cyan

# --- 0. Emplacements ---------------------------------------------------------
$Source    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Workspace = Join-Path $env:USERPROFILE "Documents\MicroStV8i\WorkSpace"

if (-not (Test-Path $Workspace)) {
    # Autre emplacement possible : installation par defaut de V8i
    $Alt = "C:\ProgramData\Bentley\MicroStation V8i (SELECTseries)\WorkSpace"
    if (Test-Path $Alt) { $Workspace = $Alt }
    else {
        Write-Host "ERREUR : workspace MicroStation introuvable." -ForegroundColor Red
        Write-Host "Cherche : $Workspace"
        Write-Host "Modifiez la variable `$Workspace en tete de script si votre"
        Write-Host "workspace est ailleurs, puis relancez."
        exit 1
    }
}
Write-Host "Workspace : $Workspace"

# --- 1. Copier le projet VBA -------------------------------------------------
$Mvba = Join-Path $Source "InterpolationTopo.mvba"
$VbaDir = Join-Path $Workspace "Standards\vba"
if (Test-Path $Mvba) {
    New-Item -ItemType Directory -Force $VbaDir | Out-Null
    Copy-Item $Mvba $VbaDir -Force
    Write-Host "[OK] InterpolationTopo.mvba copie vers $VbaDir" -ForegroundColor Green
} else {
    Write-Host "[!] InterpolationTopo.mvba absent du dossier d'installation :" -ForegroundColor Yellow
    Write-Host "    le projet VBA devra etre installe manuellement (voir README)."
}

# --- 2. Copier la boite a outils (.dgnlib) -----------------------------------
$Dgnlib = Join-Path $Source "MesMacros.dgnlib"
$GuiDir = Join-Path $Workspace "System\GUI"
if (Test-Path $Dgnlib) {
    if (-not (Test-Path $GuiDir)) {
        Write-Host "[!] Dossier System\GUI introuvable : dgnlib non copiee." -ForegroundColor Yellow
    } else {
        Copy-Item $Dgnlib $GuiDir -Force
        Write-Host "[OK] MesMacros.dgnlib (ToolBox) copiee vers $GuiDir" -ForegroundColor Green
    }
} else {
    Write-Host "[!] MesMacros.dgnlib absente : pas de ToolBox installee." -ForegroundColor Yellow
}

# --- 3. Configurer le chargement automatique (fichier .ucf utilisateur) ------
$UsersDir = Join-Path $Workspace "Users"
$Ucf = $null
if (Test-Path $UsersDir) {
    $UcfFiles = @(Get-ChildItem $UsersDir -Filter *.ucf -File)
    if ($UcfFiles.Count -ge 1) { $Ucf = $UcfFiles[0].FullName }
}
if ($null -eq $Ucf) {
    Write-Host "[!] Aucun fichier .ucf trouve dans $UsersDir" -ForegroundColor Yellow
    Write-Host "    Configurez l'autoload dans MicroStation : Utilities > Macros >"
    Write-Host "    Project Manager > clic droit sur InterpolationTopo > Autoload"
} else {
    $Contenu = Get-Content $Ucf -Raw
    $Lignes = @(
        "# --- Interpolation Topo (ajoute par install.ps1) ---",
        "MS_VBASEARCHDIRECTORIES < $VbaDir\",
        "MS_VBAAUTOLOADPROJECTS > InterpolationTopo"
    )
    if ($Contenu -match "MS_VBAAUTOLOADPROJECTS\s*>?\s*InterpolationTopo") {
        Write-Host "[OK] Autoload deja configure dans $([IO.Path]::GetFileName($Ucf))" -ForegroundColor Green
    } else {
        Add-Content -Path $Ucf -Value ("`r`n" + ($Lignes -join "`r`n") + "`r`n") -Encoding ASCII
        Write-Host "[OK] Autoload configure dans $([IO.Path]::GetFileName($Ucf))" -ForegroundColor Green
    }
}

# --- 4. Recapitulatif --------------------------------------------------------
Write-Host ""
Write-Host "=== Installation terminee ===" -ForegroundColor Cyan
Write-Host "1. (Re)demarrez MicroStation"
Write-Host "2. La ToolBox : Workspace > Customize doit lister MesMacros.dgnlib ;"
Write-Host "   ouvrez la ToolBox via clic droit > Open si elle n'apparait pas."
Write-Host "3. Test key-in (F9) : vba run [InterpolationTopo]InterpolerPoint"
Write-Host ""
Read-Host "Appuyez sur Entree pour fermer"

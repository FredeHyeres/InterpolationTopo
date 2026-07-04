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

# --- 1. Installer le projet VBA (Default MVBA) -------------------------------
# La macro vit dans le Default MVBA : charge automatiquement par MicroStation,
# aucun autoload a configurer, key-in sans crochets (vba run InterpolerPoint).
# Le Default.mvba est PERSONNEL et lie au projet MicroStation actif
# (WorkSpace\Projects\<projet>\vba\Default.mvba) : on ne le copie que la ou il
# n'existe pas encore, pour ne jamais ecraser les macros de l'utilisateur.
$Mvba = Join-Path $Source "Default.mvba"
$ProjectsDir = Join-Path $Workspace "Projects"
if (-not (Test-Path $Mvba)) {
    Write-Host "[!] Default.mvba absent du dossier d'installation :" -ForegroundColor Yellow
    Write-Host "    importez les 4 fichiers de src\ dans votre Default MVBA (voir README)."
} elseif (-not (Test-Path $ProjectsDir)) {
    Write-Host "[!] Dossier Projects introuvable : macro non installee." -ForegroundColor Yellow
} else {
    foreach ($Proj in Get-ChildItem $ProjectsDir -Directory) {
        $VbaDir = Join-Path $Proj.FullName "vba"
        $Cible = Join-Path $VbaDir "Default.mvba"
        if (Test-Path $Cible) {
            Write-Host "[!] $($Proj.Name) : Default.mvba existe deja - non remplace." -ForegroundColor Yellow
            Write-Host "    (pour ne pas ecraser vos macros existantes ; importez les"
            Write-Host "     fichiers de src\ dans ce Default MVBA via Alt+F11 > Ctrl+M)"
        } else {
            New-Item -ItemType Directory -Force $VbaDir | Out-Null
            Copy-Item $Mvba $Cible
            Write-Host "[OK] Macro installee : $Cible" -ForegroundColor Green
        }
    }
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

# --- 3. Recapitulatif --------------------------------------------------------
# Pas d'autoload a configurer : le Default MVBA est charge d'office.
Write-Host ""
Write-Host "=== Installation terminee ===" -ForegroundColor Cyan
Write-Host "1. (Re)demarrez MicroStation"
Write-Host "2. La ToolBox : Workspace > Customize doit lister MesMacros.dgnlib ;"
Write-Host "   ouvrez la ToolBox via clic droit > Open si elle n'apparait pas."
Write-Host "3. Test key-in (F9) : vba run InterpolerPoint"
Write-Host ""
Read-Host "Appuyez sur Entree pour fermer"

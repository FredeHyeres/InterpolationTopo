# Genere un couple .frm/.frx valide pour frmInterpolation via l'editeur VBA d'Excel.
# 1. Active temporairement AccessVBOM (restaure a la fin)
# 2. Cree un UserForm vide nomme frmInterpolation, injecte le code, exporte
$ErrorActionPreference = "Stop"

$srcFrm = "C:\Users\Fred\Documents\My Documents\Prog\Microstation_Nath\src\v2\frmInterpolation.frm"
$outDir = "C:\Users\Fred\Documents\My Documents\Prog\Microstation_Nath\src\v2"

# --- Extraire le code : tout apres l'en-tete designer, sans les lignes Attribute ---
$lines = [IO.File]::ReadAllLines($srcFrm)
$start = 0
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "^Attribute VB_Exposed") { $start = $i + 1; break }
}
$codeLines = $lines[$start..($lines.Count - 1)] | Where-Object { $_ -notmatch "^Attribute " }
$code = ($codeLines -join "`r`n")

# --- Registre : AccessVBOM ---
$regPath = "HKCU:\Software\Microsoft\Office\12.0\Excel\Security"
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
$old = $null
try { $old = (Get-ItemProperty $regPath -Name AccessVBOM -ErrorAction Stop).AccessVBOM } catch {}
Set-ItemProperty $regPath -Name AccessVBOM -Value 1 -Type DWord
Write-Output "AccessVBOM : ancienne valeur = $old, mise a 1"

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    $vbp = $wb.VBProject
    $comp = $vbp.VBComponents.Add(3)   # 3 = vbext_ct_MSForm
    $comp.Name = "frmInterpolation"
    $comp.CodeModule.AddFromString($code) | Out-Null

    $comp.Export("$outDir\frmInterpolation_export.frm")
    Write-Output "Export OK"

    $wb.Close($false)
} finally {
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    # Restaurer le registre
    if ($null -eq $old) { Remove-ItemProperty $regPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty $regPath -Name AccessVBOM -Value $old -Type DWord }
    Write-Output "AccessVBOM restaure ($old)"
}

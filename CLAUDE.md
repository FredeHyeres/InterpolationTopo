# Interpolation Topo - MicroStation V8i VBA

## Encodage obligatoire des fichiers VBA

Apres chaque creation ou modification de fichier `.cls`, `.bas` ou `.frm`,
normaliser en **CRLF + ANSI (Windows-1252)** avec :

```powershell
$t = [IO.File]::ReadAllText($path)
$t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
[IO.File]::WriteAllText($path, $t, [Text.Encoding]::GetEncoding(1252))
```

Raison : le Write tool produit du LF/UTF-8. MicroStation importe les `.cls`
en LF comme des modules standard au lieu de classes, ce qui casse `Implements`
et `New`.

## Formulaire (FRM/FRX)

- Le formulaire `frmInterpolation` est construit entierement au runtime
  dans `ConstruireControles` : le `.frx` est un blob binaire du formulaire vide.
- Si seul le **code** du `.frm` change (ajout de controles dans
  `ConstruireControles`, nouveaux handlers) : le `.frx` existant reste valide,
  pas besoin de le regenerer.
- Si les **proprietes du designer** changent (en-tete `Begin...End` du `.frm`) :
  regenerer le couple `.frm`/`.frx` via `scripts\export_frm_via_excel.ps1`.

## Lancement

Key-in MicroStation : `vba run [InterpolationTopoV2]InterpolerPoint`

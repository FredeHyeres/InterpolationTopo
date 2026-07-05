# Interpolation Topo — Version 2

## Objectif

Refonte architecturale de la V1 en vue d'intégrer une boîte de dialogue modeless
(équivalent VBA du Tool Settings MicroStation) et d'isoler chaque responsabilité
dans une classe dédiée.

L'algorithme de calcul et la logique de sélection sont identiques à la V1.

---

## Fichiers

| Fichier | Type | Rôle |
|---|---|---|
| `InterpolationTopoV2.bas` | Module | Point d'entrée, globals propres, `TrouverTexteProche` |
| `CMstSettings.cls` | Classe | Agrégateur unique des paramètres de la commande |
| `CSymboTexte.cls` | Classe | Attributs du texte (niveau, couleur, décimales, séparateur) |
| `CSymboCercle.cls` | Classe | Attributs du cercle (rayon, couleur, niveau) |
| `CPointRef.cls` | Classe | Point de référence topo : position + altitude + validité |
| `CInterpolation.cls` | Classe | Calcul pur : parsing, projection, interpolation, pente, gisement |
| `CMoteurGraphique.cls` | Classe | Création graphique : cercle + texte, aperçu dynamique |
| `CSelectP1.cls` | Classe | État commande : sélection du point P1 |
| `CSelectP2.cls` | Classe | État commande : sélection du point P2 |
| `CPlacerPoint.cls` | Classe | État commande : placement dynamique + création |
| `frmInterpolation.frm` | Formulaire | Dialog modeless (Tool Settings) : paramètres + état. Contrôles créés par code au chargement |
| `frmInterpolation.frx` | Binaire | Blob designer du formulaire vide (généré par l'éditeur VBA d'Excel) — **à importer avec le `.frm`, garder les deux fichiers dans le même dossier** |

---

## Architecture

```
InterpolerPoint()
      │
      ├─── CMstSettings (g_oSettings)
      │         ├── CSymboTexte  (chargé depuis P1)
      │         └── CSymboCercle (saisi via formulaire)
      │
      ├─── CPointRef × 2 (g_oP1, g_oP2)
      ├─── CInterpolation (g_oCalc)
      ├─── CMoteurGraphique (g_oMoteur)
      │
      ├─── frmInterpolation  ←→  g_oSettings (lecture/écriture)
      │
      └─── CommandState
                ├── CSelectP1  →  CSelectP2  →  CPlacerPoint
                └── (Reset remonte d'un cran, second Reset quitte)
```

---

## Ce qui change par rapport à la V1

### V1 — globals spaghetti

```vba
Public g_oP1        As Point3d
Public g_dZ1        As Double
Public g_sDecSep    As String
Public g_nDecimals  As Integer
Public g_oTextTemplate As TextElement
Public g_dCircRayon As Double
Public g_nCircColor As Long
Public g_sCircLevel As String
```

Tous les modules et classes accédaient directement à ces variables globales.
Les paramètres étaient saisis via `InputBox` (bloquant).

### V2 — classes à responsabilité unique

```vba
Public g_oSettings  As CMstSettings      ' tous les paramètres
Public g_oP1        As CPointRef         ' point 1 (position + altitude)
Public g_oP2        As CPointRef         ' point 2
Public g_oCalc      As CInterpolation    ' calcul pur
Public g_oMoteur    As CMoteurGraphique  ' création graphique
```

Les paramètres sont modifiables en temps réel via `frmInterpolation` (modeless).

---

## Séparation des responsabilités

| Domaine | Classe | Connaît MST API ? |
|---|---|---|
| Paramètres | `CMstSettings`, `CSymboTexte`, `CSymboCercle` | Non |
| Données géo | `CPointRef` | Non (sauf `Point3d` UDT) |
| Calcul | `CInterpolation` | Non |
| Graphique | `CMoteurGraphique` | Oui |
| Commande | `CSelectP1/2`, `CPlacerPoint` | Oui (`IPrimitiveCommandEvents`) |
| UI | `frmInterpolation` | Oui (lecture niveaux DGN) |

---

## Formulaire modeless

`frmInterpolation` s'affiche dès le lancement de la commande et reste visible.

**Paramètres modifiables en cours de commande :**
- Diamètre du cercle
- Couleur du cercle (index MicroStation 0–255)
- Niveau du cercle (liste des niveaux du DGN, vide = niveau du texte P1)
- Texte altitude : case « Mêmes attributs que le texte P1 » (cochée par défaut).
  Décochée, la couleur et le niveau du texte créé deviennent choisissables
  (vide = niveau du texte P1) ; les autres propriétés (police, taille, style,
  justification) restent héritées du texte P1. En mode « mêmes attributs »,
  les champs grisés affichent les valeurs héritées après sélection de P1.

**État affiché (lecture seule, mis à jour par les classes de commande) :**
- P1 : altitude sélectionnée
- P2 : altitude sélectionnée
- Pente (%) et gisement (gon) du segment

**Construction au runtime :** les contrôles sont créés par code dans
`ConstruireControles` (appelé par `UserForm_Initialize`). Le couple
`.frm`/`.frx` ne décrit que le formulaire vide (le `.frx` a été généré par
l'éditeur VBA d'Excel, seul moyen de produire un blob designer valide) ; tout
le contenu visible vient du code, donc versionnable en texte. La validation
des saisies se fait en continu (événement `Change`) : une valeur invalide est
ignorée, Entrée reformate le champ. Fermer la fenêtre par la croix termine
proprement la commande sans décharger le formulaire.

> **Note :** Il s'agit d'un `UserForm` VBA modeless, pas du Tool Settings natif
> MicroStation. Le Tool Settings natif n'est accessible qu'en MDL (C++).

---

## Installation

1. MicroStation → Utilitaires → Macros → Gestionnaire de projets VBA → Nouveau projet
2. Alt+F11 pour ouvrir l'éditeur VBA
3. Ctrl+M → sélectionner tous les fichiers de `src\v2\` en une seule fois
4. Vérifier que les `.cls` apparaissent sous **Modules de classe** (et non sous Modules)

> **Important :** L'éditeur VBA attend des fichiers **ANSI sans BOM UTF-8** et
> en fins de ligne **CRLF**. Si les `.cls` arrivent comme Modules standard ou
> avec des caractères parasites (`ï»¿`) en tête, reconvertir avec PowerShell :

```powershell
Get-ChildItem src\v2\* -Include *.bas,*.cls,*.frm | ForEach-Object {
    $t = [IO.File]::ReadAllText($_.FullName)
    $t = $t.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($_.FullName, $t, [Text.Encoding]::ASCII)
}
```

> **Import du formulaire :** `frmInterpolation.frm` exige la présence de
> `frmInterpolation.frx` dans le même dossier (référencé par la ligne
> `OleObjectBlob`). Un `.frm` seul, sans `.frx` ou avec un GUID designer
> incorrect, est refusé avec « Impossible de charger... ». En dernier recours :
> Insertion → UserForm, nommer le formulaire `frmInterpolation`, puis coller le
> contenu du `.frm` à partir de `Option Explicit` dans son module de code —
> les contrôles étant créés au runtime, aucun dessin manuel n'est nécessaire.

---

## Lancement

```
vba run [InterpolationTopoV2]InterpolerPoint
```

---

## Points à faire évoluer (futures versions)

- Persistance des paramètres entre deux sessions (fichier INI ou registre)
- Tolérance de recherche des textes configurable dans le formulaire
- Support multi-droites (mémoriser plusieurs segments P1-P2)
- Export des points créés (CSV, rapport)

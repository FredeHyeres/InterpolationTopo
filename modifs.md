# Interpolation Topo V2 - Etat des lieux

## Architecture

Macro VBA pour MicroStation V8i SS3. Interpole l'altitude d'un point sur la
droite reliant deux points d'altitude connus (P1, P2), avec apercu dynamique,
puis cree un texte altitude et un cercle (ou clone l'element source).

Lancement : `vba run [InterpolationTopoV2]InterpolerPoint`

### Fichiers (src/v2/)

| Fichier | Role |
|---|---|
| `InterpolationTopoV2.bas` | Point d'entree, globals, `TrouverTexteProche`, `TrouverAltitudeProche`, utilitaires scan |
| `CSelectP1.cls` | Etat commande : selection du point de reference P1 |
| `CSelectP2.cls` | Etat commande : selection du point de reference P2 |
| `CPlacerPoint.cls` | Etat commande : placement en 2 etapes (projection + decalage perpendiculaire) |
| `CInterpolation.cls` | Moteur de calcul pur (parsing, projection, interpolation, formatage) |
| `CMoteurGraphique.cls` | Creation des elements MicroStation (texte, cercle, clonage cellule/tag) + apercu dynamique |
| `CMstSettings.cls` | Agregateur des parametres (symbologie texte/cercle, tolerance, pente decalage) |
| `CSymboTexte.cls` | Attributs de symbologie du texte altitude |
| `CSymboCercle.cls` | Attributs de symbologie du cercle |
| `CPointRef.cls` | Point de reference (position XY + altitude) |
| `frmInterpolation.frm` | Formulaire modeless "Tool Settings" (controles crees au runtime) |
| `frmInterpolation.frx` | Binaire associe au formulaire (genere par `scripts/export_frm_via_excel.ps1`) |

### Scripts

| Fichier | Role |
|---|---|
| `scripts/export_frm_via_excel.ps1` | Genere un couple .frm/.frx valide via l'editeur VBA d'Excel |

## Fonctionnalites implementees

### Detection des altitudes sources

La fonction `TrouverTexteProche` (+ wrapper `TrouverAltitudeProche`) detecte
trois types d'elements portant une altitude numerique :

1. **Textes simples** (`msdElementTypeText`) - cas de base
2. **Cellules** (`msdElementTypeCellHeader`) - cherche un texte ou un tag
   numerique parmi les sous-elements de la cellule
3. **Tags** (`msdElementTypeTag`) - tags isoles avec valeur numerique

Les elements sur un **niveau gele** (non affiche dans la vue active) sont
ignores (`EstSurNiveauGele` via `Level.IsDisplayedInView`).

### Creation du point interpole

Trois strategies selon le type de l'element source (dans `CMoteurGraphique`) :

| Source | Methode | Detail |
|---|---|---|
| Texte simple | `CreerTexteEtCercle` | Cree un TextElement (modele P1) + EllipseElement |
| Cellule | `ClonerCellule` | Clone la cellule entiere, deplace, modifie le texte ou tag interne |
| Tag isole | `ClonerTag` | Clone le tag seul (apparait "orphelin" - probleme connu) |

### Placement en 2 etapes (`CPlacerPoint`)

1. **Etape 1** : le curseur suit la droite P1-P2, l'altitude est interpolee
   en temps reel. Un clic fige l'altitude et le point projete.
2. **Etape 2** : une perpendiculaire a P1-P2 apparait depuis le point projete.
   Le curseur se deplace le long de cette perpendiculaire. Un clic cree le
   point a la position decalee.

### Pente transversale

Option dans le formulaire (cadre "Pente decalage") :
- Checkbox "Appliquer pente transversale" + champ pente en %
- Quand active, l'altitude du point decale est corrigee :
  `Z_final = Z_interpole + decalage * pente / 100`
- Ne s'applique qu'a l'etape 2 (apres le premier clic)

### Formulaire (frmInterpolation)

Controles crees au runtime (pas de dependance au designer) :

- **Cadre Cercle** : diametre, couleur, niveau
- **Cadre Texte altitude** : checkbox "memes attributs que P1", couleur, niveau
- **Cadre Pente decalage** : checkbox activer/desactiver, pente en %
- **Cadre Etat** : affiche P1, P2, infos segment (pente, gisement)

## Problemes connus

### Tags isoles (non resolus)

Quand l'altitude source est un tag **isole** (pas dans une cellule), le clonage
du tag seul produit un element "orphelin" affiche en pointille. La recherche du
parent via `ParentID` n'a pas fonctionne. La detection des tags **a l'interieur
des cellules** fonctionne correctement (via `ExtraireAltitudeDeCellule`).

Piste : certains tags sont des sous-elements de cellules mais le scan
`msdElementTypeTag` les trouve en premier (avant le scan `msdElementTypeCellHeader`).
La fonction `ExtraireAltitudeDeCellule` cherche deja les tags dans les cellules,
mais il faut que la cellule parente soit effectivement scannee.

### Modele texte pour tags

Quand l'altitude source est un tag (meme dans une cellule), les proprietes de
texte (police, taille) ne sont pas copiees sur le TextElement modele car
`TagElement` n'expose pas `TextStyle` et les proprietes individuelles (`Font`,
`Height`, `Width`) ne se transferent pas correctement sur un `TextElement` cree
par `CreateTextElement1`. Le clonage de cellule contourne ce probleme.

## Historique des modifications (session courante)

1. **Support des cellules** : scan de `msdElementTypeCellHeader`, extraction du
   texte numerique parmi les sous-elements, distance calculee depuis l'origin
   de la cellule

2. **Filtrage des niveaux geles** : fonction `EstSurNiveauGele` avec
   `Level.IsDisplayedInView`, protegee par `On Error` pour les elements sans
   niveau valide

3. **Support des tags** : scan de `msdElementTypeTag`, lecture de `.Value`,
   globales `g_oTagTrouve`, `g_sTagDefName`, `g_ptOrigineTexte`, `g_sValeurTexte`

4. **Refonte de LireAltitudeTexte** : utilise `g_ptOrigineTexte` et
   `g_sValeurTexte` au lieu de `oText.Origin` et `oText.Text` (suppression du
   parametre `TextElement`)

5. **Clonage des elements** : cellules clonees via `CopyElement` + `Move` +
   modification interne ; tags clones (fallback orphelin)

6. **Detection tags dans cellules** : `ExtraireAltitudeDeCellule` remplace
   `ExtraireTexteDeCellule`, cherche tags ET textes dans les sous-elements

7. **Placement en 2 etapes** : etape 1 = figer altitude sur P1-P2, etape 2 =
   decalage perpendiculaire avec apercu dynamique (`DessinDecalage`)

8. **Pente transversale** : parametres `dPenteDecalage` / `bAppliquerPente`
   dans `CMstSettings`, cadre dans le formulaire, application a l'etape 2

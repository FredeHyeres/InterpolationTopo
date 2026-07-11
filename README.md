# Interpolation Topo V2 — MicroStation V8i (VBA)

Macros VBA pour **MicroStation V8i SS3** dediees aux leves topographiques :
interpolation d'altitude, chemin de pente et rayonnement ponctuel.

Pensees pour les leves topographiques : gisement en **gon** (0 = nord, sens horaire),
pente en **%**, separateur decimal et nombre de decimales repris automatiquement des
textes source, cercles placables sur un niveau distinct des textes.

## Deux commandes

| Commande | Key-in | Description |
|---|---|---|
| **Interpolation** | `vba run [InterpolationTopoV2]InterpolerPoint` | Interpole l'altitude sur la droite P1-P2, avec pente transversale, DZ et indicateur de pente |
| **Interpol. Ponctuelle** | `vba run [InterpolationTopoV2]InterpolPonctuelle` | Chemin principal depuis P1 (pente/DZ cumulatifs) + rayonnement lateral depuis chaque point du chemin |

## Fonctionnalites communes

- **Sources d'altitude multiples** : textes simples, tags et cellules contenant des tags
  d'altitude (les elements non numeriques sont ignores automatiquement)
- **Lecture des references** : les altitudes dans les fichiers de reference attaches sont
  detectees et converties dans le repere du modele actif
- Le texte du point 1 sert de **modele de mise en forme** : police, taille, niveau,
  couleur et rotation sont repris pour le texte cree
- **Snap terrain optionnel** : apres selection du texte altitude, un second clic permet de
  positionner P1/P2 sur le point terrain reel (Reset = garder l'ancrage du texte)
- **Formulaire modeless** (Tool Settings) : parametres modifiables en temps reel
  pendant la commande
- **Apercu dynamique** avec 3 decimales pour le controle visuel
- **Nombre de decimales** des points crees configurable (0 a 6)
- **Positionnement 3D** : les elements crees sont places a l'altitude calculee (Z)

## Interpolation (P1-P2)

- **Projection + perpendiculaire en un seul clic** : le curseur calcule simultanement
  la projection sur P1-P2 et le decalage perpendiculaire, affichant le Z sur la droite
  (vert) et le Z final au point deporte
- **Pente transversale** : applique un delta Z proportionnel a la distance au segment
  P1-P2 (meme pente des deux cotes de la ligne de reference). Bouton **+/-** pour
  inverser la pente
- **Decalage fixe DZ** : ajoute une valeur constante a l'altitude interpolee
- **Indicateur de pente** : place un texte affichant la pente signee (+/-) et une fleche
  dirigee de P1 vers P2. Proprietes personnalisables : hauteur, largeur, couleur du texte,
  couleur et longueur de la fleche, niveau
- **Apercu dynamique** pendant le placement :
  - ligne support P1-P2 (etendue de 20 % de part et d'autre)
  - croix sur P1 et P2, cercle et texte provisoires suivant le curseur
  - barre d'etat : altitude interpolee, abscisse `t` (0 = P1, 1 = P2), pente %,
    gisement gon
  - avertissement en cas d'extrapolation hors du segment
- Creation **repetable** : plusieurs points sur la meme droite (Data = creer,
  Reset = nouvelle selection)

## Interpolation Ponctuelle (chemin + rayonnement)

- **P2 optionnel** : si selectionne, la pente P1-P2 est calculee automatiquement et
  affectee au chemin principal. Sinon, la pente est saisie manuellement dans le formulaire
- **Chemin principal** : ligne dynamique depuis le dernier point, altitude calculee avec
  pente (%) et/ou DZ cumulatifs. Ligne verte
- **Rayonnement** : apres chaque point du chemin, possibilite de placer des points
  lateraux avec pente/DZ propres au rayonnement. Ligne jaune
- **Retour au chemin** : clic droit en mode rayonnement revient au chemin depuis le
  dernier point chemin
- **Formulaire dedie** : cadres Chemin principal (pente/DZ), Rayonnement (pente/DZ),
  decimales partagees, cercle, texte, etat

## Installation automatique (recommandee)

1. Telecharger le depot :
   [**telecharger le ZIP**](https://github.com/FredeHyeres/InterpolationTopo/archive/refs/heads/main.zip)
   (ou **Code > Download ZIP**) et le decompresser
2. Double-cliquer sur **`install.cmd`**

Le script copie le projet **`Interpolation.mvba`** dans le workspace
(`Standards\vba`), installe la boite a outils (`MesMacros.dgnlib`) et configure
son chargement automatique au demarrage dans le `.ucf` utilisateur. Il ne touche
**jamais** au `Default.mvba` personnel. Relancable sans risque.

> Le script suppose le workspace dans `Documents\MicroStV8i\WorkSpace`
> (sinon, modifier la variable `$Workspace` en tete de `install.ps1`).

## Installation manuelle

1. MicroStation : *Utilitaires > Macros > Gestionnaire de projets VBA*, projet
   **Default** (ou un nouveau projet dedie)
2. Editeur VBA (`Alt+F11`) : *Fichier > Importer un fichier* (`Ctrl+M`) et importer
   tous les fichiers du dossier [`src/`](src/) :

   | Fichier | Type apres import |
   |---|---|
   | `InterpolationTopoV2.bas` | Module (globals, init commune, entree Interpolation) |
   | `InterpolPonctuelle.bas` | Module (entree Ponctuelle) |
   | `RechercheAltitude.bas` | Module (recherche des altitudes sources) |
   | `CMstSettings.cls` | Module de classe |
   | `CSymboTexte.cls` | Module de classe |
   | `CSymboCercle.cls` | Module de classe |
   | `CSymboPente.cls` | Module de classe |
   | `CParamPenteDZ.cls` | Module de classe |
   | `CPointRef.cls` | Module de classe |
   | `CInterpolation.cls` | Module de classe |
   | `CMoteurGraphique.cls` | Module de classe |
   | `CAltitudeSelection.cls` | Module de classe |
   | `CSelectP1.cls` | Module de classe (commune aux deux commandes) |
   | `CSnapP1.cls` | Module de classe (commune aux deux commandes) |
   | `CSelectP2.cls` | Module de classe (commune aux deux commandes) |
   | `CSnapP2.cls` | Module de classe (commune aux deux commandes) |
   | `CPlacerPoint.cls` | Module de classe (Interpolation) |
   | `CPlacerChemin.cls` | Module de classe (Interpolation) |
   | `CPlacerPente.cls` | Module de classe (Interpolation) |
   | `CPlacerPonctuel.cls` | Module de classe (Ponctuelle) |
   | `frmInterpolation.frm` + `.frx` | UserForm (Interpolation) |
   | `frmInterpolPonct.frm` + `.frx` | UserForm (Ponctuelle) |

3. *Debogage > Compiler*, puis enregistrer

> Les fichiers doivent conserver des fins de ligne **Windows (CRLF)** -- le
> `.gitattributes` du depot s'en charge. Avec des fins de ligne Unix (LF), l'editeur
> VBA importe les `.cls` comme modules standard et la compilation echoue.

> Les fichiers `frmInterpolation.frx` et `frmInterpolPonct.frx` doivent etre dans
> le **meme dossier** que leur `.frm` respectif lors de l'import.

## Lancement

Key-in (le projet `Interpolation.mvba` etant charge automatiquement) :

```
vba run [InterpolationTopoV2]InterpolerPoint
vba run [InterpolationTopoV2]InterpolPonctuelle
```

## Affectation a des touches de fonction

Menu *Utilitaires > Touches de fonction*, par exemple :

| Touche | Key-in |
|---|---|
| `F6` | `vba run [InterpolationTopoV2]InterpolerPoint` |
| `F7` | `vba run [InterpolationTopoV2]InterpolPonctuelle` |

Les key-in peuvent aussi etre associes a des **boutons de boite a outils** -- la ToolBox
`MesMacros.dgnlib` installee par le script contient deja ces boutons.

Guide detaille pas a pas (installation, utilisation, touche de fonction, boite a
outils, depannage) : **[Mode_Emploi_InterpolationTopo.html](Mode_Emploi_InterpolationTopo.html)**
(a ouvrir dans un navigateur).

## Utilisation en bref — Interpolation

1. Lancer la commande -- le formulaire **Interpolation Topo** s'ouvre (Tool Settings)
2. Cliquer pres du texte/tag altitude du **point 1** (modele de mise en forme)
3. **Snap terrain P1** : cliquer sur le point terrain reel (Reset = garder l'ancrage du texte)
4. Cliquer pres du texte/tag altitude du **point 2** -- pente et gisement s'affichent
5. **Snap terrain P2** : idem
6. Deplacer le curseur : projection sur P1-P2 + perpendiculaire calculees simultanement,
   Z sur la droite (vert) et Z final affiches en dynamique
7. **Data** = creer le point (repetable) -- **Reset** = nouvelle selection de points
   (retour a l'etape 2, memes parametres) -- **Reset** au choix du point 1 = quitter

## Utilisation en bref — Interpolation Ponctuelle

1. Lancer la commande -- le formulaire **Interpol. Ponctuelle** s'ouvre
2. Cliquer pres du texte/tag altitude du **point 1** (modele)
3. **Snap terrain P1**
4. **P2 optionnel** : Data = selectionner P2 (pente auto), Reset = pente manuelle
5. **Mode chemin** : ligne verte depuis P1, Z = Z precedent + pente*distance + DZ.
   Data = creer un point, Reset = nouvelle selection
6. Apres chaque point chemin : **mode rayonnement** (ligne jaune).
   Data = creer un point rayon, Reset = retour au chemin
7. Le chemin repart du dernier point chemin cree

### Indicateur de pente

Apres selection et snap de P2, le bouton **Pente + fleche** du formulaire bascule vers
le placement de l'indicateur de pente :

- Un texte affichant la pente signee (ex. `+2,50%` ou `-1,30%`)
- Une fleche dirigee de P1 vers P2

Cliquer pour placer l'indicateur (repetable). Le bouton **Interpolation** revient au
placement de points interpoles.

### Formulaire (modifiable a tout moment)

- **Cercle** : diametre, couleur (index 0-255), niveau
- **Texte altitude** : case "Memes attributs que le texte P1" (cochee par defaut).
  Decochee : couleur et niveau du texte deviennent choisissables
- **Pente decalage** : case "Appliquer pente transversale" + valeur en %.
  Bouton **+/-** pour inverser. La pente s'applique identiquement des deux cotes
  de la ligne P1-P2
- **Decalage fixe DZ** : case a cocher + valeur en unites maitre, ajoutee a
  l'altitude interpolee
- **Indicateur pente** : case "Personnaliser" + hauteur/largeur du texte, niveau,
  couleur du texte, couleur de la fleche, longueur de la fleche
- **Etat** : altitudes P1/P2, pente %, gisement gon
- **Actions** : boutons Interpolation / Pente + fleche (actifs apres snap P2)

## Architecture

Voir [src/README_V2.md](src/README_V2.md) pour le detail des classes et
de l'architecture.

## Notes de compatibilite

Teste sur MicroStation V8i SS3 (08.11.09.578). Le code utilise volontairement :

- `CommandState.StartPrimitive` (certaines installations n'exposent pas
  `StartPrimitiveCommand`)
- des `Sub` avec parametre de sortie plutot que des `Function` retournant `Point3d` /
  `Matrix3d` (refusees par certains environnements VBA)
- `CommandState.StartDynamics` dans le `Start` de la commande de placement,
  indispensable pour que l'evenement `Dynamics` (apercu) se declenche
- `Matrix3dFromAxisAndRotationAngle(2, angle)` pour les rotations (l'affectation
  directe des membres d'une `Matrix3d` imbriquee echoue silencieusement en VBA)

## Licence

Utilisation et modification libres. Fourni tel quel, sans garantie.

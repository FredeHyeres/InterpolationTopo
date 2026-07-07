# Interpolation Topo V2 — MicroStation V8i (VBA)

Macro VBA pour **MicroStation V8i SS3** qui interpole l'altitude d'un point situé sur la
droite reliant deux points d'altitude existants (textes topo ou tags/cellules), avec
**aperçu dynamique** et **formulaire modeless** (Tool Settings), puis crée le texte
altitude et un cercle au point choisi.

Pensée pour les levés topographiques : gisement en **gon** (0 = nord, sens horaire),
pente en **%**, séparateur décimal et nombre de décimales repris automatiquement des
textes source, cercles plaçables sur un niveau distinct des textes.

## Fonctionnalités

- **Sources d'altitude multiples** : textes simples, tags et cellules contenant des tags
  d'altitude (les éléments non numériques sont ignorés automatiquement)
- Le texte du point 1 sert de **modèle de mise en forme** : police, taille, niveau,
  couleur et rotation sont repris pour le texte créé
- **Formulaire modeless** (Tool Settings) : paramètres modifiables en temps réel
  pendant la commande — diamètre, couleur et niveau du cercle ; couleur et niveau du
  texte ; pente transversale
- **Pente transversale** : applique un delta Z proportionnel à la distance au segment
  P1–P2 (même pente des deux côtés de la ligne de référence). Bouton **+/-** pour
  inverser la pente
- **Aperçu dynamique** pendant le placement :
  - ligne support P1–P2 (étendue de 20 % de part et d'autre)
  - croix sur P1 et P2, cercle et texte provisoires suivant le curseur
  - barre d'état : altitude interpolée, abscisse `t` (0 = P1, 1 = P2), pente %,
    gisement gon
  - avertissement en cas d'extrapolation hors du segment
- Création **répétable** : plusieurs points sur la même droite (Data = créer,
  Reset = nouvelle sélection)

## Installation automatique (recommandée)

1. Télécharger le dépôt :
   [**⬇ télécharger le ZIP**](https://github.com/FredeHyeres/InterpolationTopo/archive/refs/heads/main.zip)
   (ou **Code > Download ZIP**) et le décompresser
2. Double-cliquer sur **`install.cmd`**

Le script copie le projet **`Interpolation.mvba`** dans le workspace
(`Standards\vba`), installe la boîte à outils (`MesMacros.dgnlib`) et configure
son chargement automatique au démarrage dans le `.ucf` utilisateur. Il ne touche
**jamais** au `Default.mvba` personnel. Relançable sans risque.

> Le script suppose le workspace dans `Documents\MicroStV8i\WorkSpace`
> (sinon, modifier la variable `$Workspace` en tête de `install.ps1`).

## Installation manuelle

1. MicroStation : *Utilitaires > Macros > Gestionnaire de projets VBA*, projet
   **Default** (ou un nouveau projet dédié)
2. Éditeur VBA (`Alt+F11`) : *Fichier > Importer un fichier* (`Ctrl+M`) et importer
   tous les fichiers du dossier [`src/v2/`](src/v2) :

   | Fichier | Type après import |
   |---|---|
   | `InterpolationTopoV2.bas` | Module |
   | `CMstSettings.cls` | Module de classe |
   | `CSymboTexte.cls` | Module de classe |
   | `CSymboCercle.cls` | Module de classe |
   | `CPointRef.cls` | Module de classe |
   | `CInterpolation.cls` | Module de classe |
   | `CMoteurGraphique.cls` | Module de classe |
   | `CSelectP1.cls` | Module de classe |
   | `CSelectP2.cls` | Module de classe |
   | `CPlacerPoint.cls` | Module de classe |
   | `frmInterpolation.frm` + `.frx` | UserForm |

3. *Débogage > Compiler*, puis enregistrer

> ⚠️ Les fichiers doivent conserver des fins de ligne **Windows (CRLF)** — le
> `.gitattributes` du dépôt s'en charge. Avec des fins de ligne Unix (LF), l'éditeur
> VBA importe les `.cls` comme modules standard et la compilation échoue.

> ⚠️ Le fichier `frmInterpolation.frx` doit être dans le **même dossier** que le
> `.frm` lors de l'import. Sans lui, l'import échoue avec « Impossible de charger... ».

## Lancement

Key-in (le projet `Interpolation.mvba` étant chargé automatiquement) :

```
vba run [InterpolationTopoV2]InterpolerPoint
```

Le key-in peut être associé à une **touche de fonction** (*Utilitaires > Touches de
fonction*) ou à un **bouton de boîte à outils** — la ToolBox `MesMacros.dgnlib`
installée par le script contient déjà ce bouton.

📖 Guide détaillé pas à pas (installation, utilisation, touche de fonction, boîte à
outils, dépannage) : **[Mode_Emploi_InterpolationTopo.html](Mode_Emploi_InterpolationTopo.html)**
(à ouvrir dans un navigateur).

## Utilisation en bref

1. Lancer la commande → le formulaire **Interpolation Topo** s'ouvre (Tool Settings)
2. Cliquer près du texte/tag altitude du **point 1** (modèle de mise en forme)
3. Cliquer près du texte/tag altitude du **point 2** → pente et gisement s'affichent
4. Déplacer le curseur : la position est projetée sur la droite P1–P2, l'altitude
   interpolée s'affiche en dynamique
5. **Data** = créer le point (répétable) · **Reset** = nouvelle sélection de points
   (retour à l'étape 2, mêmes paramètres) · **Reset** au choix du point 1 = quitter

### Formulaire (modifiable à tout moment)

- **Cercle** : diamètre, couleur (index 0–255), niveau
- **Texte altitude** : case « Mêmes attributs que le texte P1 » (cochée par défaut).
  Décochée → couleur et niveau du texte deviennent choisissables
- **Pente décalage** : case « Appliquer pente transversale » + valeur en %.
  Bouton **+/-** pour inverser. La pente s'applique identiquement des deux côtés
  de la ligne P1–P2

## Architecture

Voir [src/v2/README_V2.md](src/v2/README_V2.md) pour le détail des classes et
de l'architecture.

## Notes de compatibilité

Testé sur MicroStation V8i SS3 (08.11.09.578). Le code utilise volontairement :

- `CommandState.StartPrimitive` (certaines installations n'exposent pas
  `StartPrimitiveCommand`)
- des `Sub` avec paramètre de sortie plutôt que des `Function` retournant `Point3d` /
  `Matrix3d` (refusées par certains environnements VBA)
- `CommandState.StartDynamics` dans le `Start` de la commande de placement,
  indispensable pour que l'événement `Dynamics` (aperçu) se déclenche

## Licence

Utilisation et modification libres. Fourni tel quel, sans garantie.
